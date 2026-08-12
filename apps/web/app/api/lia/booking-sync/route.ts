import process from "node:process";
import prisma from "@calcom/prisma";
import type { Prisma } from "@calcom/prisma/client";
import { BookingStatus, WebhookTriggerEvents } from "@calcom/prisma/enums";
import { defaultResponderForAppDir } from "app/api/defaultResponderForAppDir";
import type { NextRequest } from "next/server";
import { NextResponse } from "next/server";
import z from "zod";

const bookingSyncSchema = z.object({
  subscriberUrl: z
    .string()
    .url()
    .refine((value) => new URL(value).protocol === "https:", {
      message: "subscriberUrl must use HTTPS",
    }),
  webhookSecret: z.string().min(16),
  updatedAfter: z.string().datetime(),
  cursor: z.object({ updatedAt: z.string().datetime(), id: z.number().int().nonnegative() }).optional(),
  limit: z.number().int().min(1).max(100).default(100),
});

const bookingSelect = {
  id: true,
  uid: true,
  createdAt: true,
  updatedAt: true,
  startTime: true,
  endTime: true,
  status: true,
  location: true,
  responses: true,
  metadata: true,
  fromReschedule: true,
  userPrimaryEmail: true,
  user: { select: { name: true, email: true, timeZone: true, metadata: true } },
  attendees: { select: { name: true, email: true, timeZone: true, phoneNumber: true } },
  references: { select: { type: true, meetingUrl: true } },
  eventType: { select: { metadata: true } },
} satisfies Prisma.BookingSelect;

type BookingSnapshot = Prisma.BookingGetPayload<{ select: typeof bookingSelect }>;

async function handler(req: NextRequest) {
  const configuredSecret = process.env.LIA_INTERNAL_SECRET ?? process.env.NEXTAUTH_SECRET;
  const configuredSubscriberUrl = process.env.LIA_BOOKING_WEBHOOK_URL;
  const organizerDomain = process.env.LIA_BOOKING_ORGANIZER_DOMAIN?.trim().toLowerCase();
  if (!configuredSecret || !configuredSubscriberUrl || !organizerDomain) {
    return NextResponse.json({ message: "LIA booking sync is not configured" }, { status: 503 });
  }
  if (req.headers.get("x-lia-internal-secret") !== configuredSecret) {
    return NextResponse.json({ message: "Invalid LIA booking sync secret" }, { status: 403 });
  }

  const parsed = bookingSyncSchema.safeParse(await req.json().catch(() => null));
  if (!parsed.success) {
    return NextResponse.json({ message: parsed.error.message }, { status: 400 });
  }
  const input = parsed.data;
  if (input.subscriberUrl !== configuredSubscriberUrl) {
    return NextResponse.json({ message: "Unapproved LIA booking subscriber" }, { status: 403 });
  }
  if (!/^(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$/.test(organizerDomain)) {
    return NextResponse.json({ message: "LIA booking organizer domain is invalid" }, { status: 503 });
  }
  const eventTriggers = [
    WebhookTriggerEvents.BOOKING_CREATED,
    WebhookTriggerEvents.BOOKING_RESCHEDULED,
    WebhookTriggerEvents.BOOKING_CANCELLED,
  ];
  const users = await prisma.user.findMany({
    where: { email: { endsWith: `@${organizerDomain}`, mode: "insensitive" } },
    select: { id: true },
  });
  const webhookIds = await Promise.all(
    users.map(async (user) => {
      const existing = await prisma.webhook.findUnique({
        where: { courseIdentifier: { userId: user.id, subscriberUrl: configuredSubscriberUrl } },
        select: { id: true },
      });
      const id = existing?.id ?? `lia-wn-sales-user-${user.id}`;
      const webhook = await prisma.webhook.upsert({
        where: { id },
        create: {
          id,
          userId: user.id,
          subscriberUrl: configuredSubscriberUrl,
          secret: input.webhookSecret,
          platform: false,
          active: true,
          eventTriggers,
        },
        update: {
          subscriberUrl: configuredSubscriberUrl,
          secret: input.webhookSecret,
          platform: false,
          active: true,
          eventTriggers,
        },
        select: { id: true },
      });
      return webhook.id;
    })
  );
  await prisma.webhook.updateMany({
    where: {
      id: { startsWith: "lia-wn-sales-user-", notIn: webhookIds },
      active: true,
    },
    data: { active: false },
  });

  const lowerBound = input.cursor ? new Date(input.cursor.updatedAt) : new Date(input.updatedAfter);
  const bookings = await prisma.booking.findMany({
    where: {
      updatedAt: { gte: lowerBound },
      user: { email: { endsWith: `@${organizerDomain}`, mode: "insensitive" } },
      ...(input.cursor
        ? {
            OR: [{ updatedAt: { gt: lowerBound } }, { updatedAt: lowerBound, id: { gt: input.cursor.id } }],
          }
        : {}),
    },
    orderBy: [{ updatedAt: "asc" }, { id: "asc" }],
    take: input.limit,
    select: bookingSelect,
  });
  const last = bookings.at(-1);

  return NextResponse.json({
    ok: true,
    webhookIds,
    events: bookings.map(toWebhookSnapshot),
    nextCursor: last
      ? { updatedAt: bookingUpdatedAt(last).toISOString(), id: last.id }
      : (input.cursor ?? null),
    hasMore: bookings.length === input.limit,
  });
}

export const POST = defaultResponderForAppDir(handler);

function bookingUpdatedAt(booking: BookingSnapshot): Date {
  return booking.updatedAt ?? booking.createdAt;
}

function toWebhookSnapshot(booking: BookingSnapshot) {
  const triggerEvent =
    booking.status === BookingStatus.CANCELLED
      ? WebhookTriggerEvents.BOOKING_CANCELLED
      : booking.fromReschedule
        ? WebhookTriggerEvents.BOOKING_RESCHEDULED
        : WebhookTriggerEvents.BOOKING_CREATED;
  const organizerEmail = booking.userPrimaryEmail ?? booking.user?.email;
  const organizer = organizerEmail
    ? {
        name: booking.user?.name ?? organizerEmail,
        email: organizerEmail,
        timeZone: booking.user?.timeZone ?? "UTC",
      }
    : undefined;
  const metadata = {
    ...jsonObject(booking.eventType?.metadata),
    ...jsonObject(booking.user?.metadata),
    ...jsonObject(booking.metadata),
  };
  return {
    triggerEvent,
    createdAt: bookingUpdatedAt(booking).toISOString(),
    payload: {
      uid: booking.uid,
      startTime: booking.startTime.toISOString(),
      endTime: booking.endTime.toISOString(),
      timeZone: booking.user?.timeZone ?? booking.attendees[0]?.timeZone ?? "UTC",
      attendees: booking.attendees,
      ...(organizer ? { organizer } : {}),
      ...(booking.location ? { location: booking.location } : {}),
      responses: booking.responses,
      metadata,
      conferenceData: booking.references.flatMap((reference) =>
        reference.meetingUrl ? [{ type: reference.type, meetingUrl: reference.meetingUrl }] : []
      ),
      ...(booking.fromReschedule ? { rescheduleUid: booking.fromReschedule } : {}),
    },
  };
}

function jsonObject(value: Prisma.JsonValue | null | undefined): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : {};
}
