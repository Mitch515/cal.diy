import { timingSafeEqual } from "node:crypto";
import process from "node:process";
import { MSTeamsLocationType } from "@calcom/app-store/constants";
import { getRegularBookingService } from "@calcom/features/bookings/di/RegularBookingService.container";
import {
  submitWncConfirmation,
  wncConfirmationConfigured,
  wncMailStateSchema,
} from "@calcom/features/bookings/lib/wnc-confirmation";
import { getAvailableSlotsService } from "@calcom/features/di/containers/AvailableSlots";
import { checkRateLimitAndThrowError } from "@calcom/lib/checkRateLimitAndThrowError";
import { ErrorCode } from "@calcom/lib/errorCodes";
import { ErrorWithCode } from "@calcom/lib/errors";
import prisma from "@calcom/prisma";
import { BookingStatus, CreationSource } from "@calcom/prisma/enums";
import { defaultResponderForAppDir } from "app/api/defaultResponderForAppDir";
import { type NextRequest, NextResponse } from "next/server";
import { z } from "zod";

const requestId = z.string().regex(/^[a-f0-9]{64}$/);
const attendee = z.object({
  name: z.string().trim().min(1).max(200),
  email: z.string().email(),
  timeZone: z.string().min(1),
});
const eventFields = {
  eventTypeId: z.number().int().positive(),
  durationMinutes: z.union([z.literal(15), z.literal(60)]),
};
const schema = z.discriminatedUnion("action", [
  z.object({ action: z.literal("status"), requestId }),
  z.object({ action: z.literal("confirmation"), requestId }),
  z.object({ action: z.literal("details"), ...eventFields }),
  z.object({
    action: z.literal("slots"),
    ...eventFields,
    start: z.string().datetime(),
    end: z.string().datetime(),
    timeZone: z.string().min(1),
  }),
  z.object({
    action: z.literal("book"),
    ...eventFields,
    requestId,
    start: z.string().datetime(),
    attendee,
    setterEmail: z
      .string()
      .email()
      .refine((email) => email.endsWith("@getwealthnavigator.com")),
  }),
]);
const metadataSchema = z.object({
  wncAttendeeEmail: z.string().email(),
  wncAttendeeTimeZone: z.string(),
  wncSetterEmail: z.string().email(),
  wncConfirmation: wncMailStateSchema.optional(),
});
const slotSchema = z.record(z.array(z.object({ time: z.string() })));

function authorized(secret: string, supplied: string | null) {
  if (!supplied) return false;
  const expected = Buffer.from(secret);
  const actual = Buffer.from(supplied);
  return expected.length === actual.length && timingSafeEqual(expected, actual);
}
function configured(name: string) {
  const value = Reflect.get(process.env, name);
  return typeof value === "string" ? value : undefined;
}
async function snapshot(id: string) {
  // Cal owns the slot collision key and clears it on cancellation. Keep our request identity separate.
  const bookings = await prisma.booking.findMany({
    where: {
      OR: [
        { idempotencyKey: `wnc:${id}` },
        { metadata: { path: ["wncRequestId"], equals: id }, fromReschedule: null },
      ],
    },
    take: 2,
    select: {
      uid: true,
      title: true,
      eventTypeId: true,
      startTime: true,
      endTime: true,
      createdAt: true,
      status: true,
      rescheduled: true,
      metadata: true,
      user: { select: { name: true, email: true } },
      attendees: { select: { name: true, email: true } },
      references: { select: { type: true, meetingUrl: true } },
    },
  });
  if (bookings.length > 1)
    throw new ErrorWithCode(ErrorCode.InternalServerError, "Ambiguous WNC booking request");
  const booking = bookings[0];
  if (!booking) return null;
  const metadata = metadataSchema.parse(booking.metadata);
  const booker = booking.attendees.find(
    (item) => item.email.toLowerCase() === metadata.wncAttendeeEmail.toLowerCase()
  );
  if (!booker || !booking.user || !booking.eventTypeId)
    throw new ErrorWithCode(ErrorCode.InternalServerError, "Incomplete WNC booking record");
  const meetingUrl = booking.references
    .map((reference) => reference.meetingUrl)
    .find((url) => {
      if (!url) return false;
      try {
        const parsed = new URL(url);
        return (
          parsed.protocol === "https:" &&
          parsed.hostname === "teams.microsoft.com" &&
          parsed.pathname.startsWith("/l/meetup-join/")
        );
      } catch {
        return false;
      }
    });
  return {
    uid: booking.uid,
    title: booking.title,
    eventTypeId: booking.eventTypeId,
    start: booking.startTime.toISOString(),
    end: booking.endTime.toISOString(),
    createdAt: booking.createdAt.toISOString(),
    attendee: { ...booker, timeZone: metadata.wncAttendeeTimeZone },
    organizer: { name: booking.user.name ?? booking.user.email, email: booking.user.email },
    participants: booking.attendees.filter(
      (person) => person.email.toLowerCase() !== booker.email.toLowerCase()
    ),
    guests: booking.attendees
      .filter((person) => person.email.toLowerCase() !== booker.email.toLowerCase())
      .map((person) => person.email),
    setterEmail: metadata.wncSetterEmail,
    confirmationStatus: metadata.wncConfirmation?.state ?? "pending",
    ...(meetingUrl ? { meetingUrl } : {}),
    status: booking.rescheduled
      ? "rescheduled"
      : booking.status === BookingStatus.CANCELLED
        ? "cancelled"
        : booking.status === BookingStatus.ACCEPTED
          ? "confirmed"
          : "pending",
    calendarStatus:
      booking.references.some((reference) => reference.type === "office365_calendar") && meetingUrl
        ? "created"
        : "pending",
  };
}

async function handler(req: NextRequest) {
  const secret = configured("LIA_INTERNAL_SECRET");
  if (!secret) return NextResponse.json({ error: "Internal booking unavailable" }, { status: 503 });
  if (!authorized(secret, req.headers.get("x-lia-internal-secret")))
    return NextResponse.json({ error: "Forbidden" }, { status: 403 });
  const parsed = schema.safeParse(await req.json().catch(() => null));
  if (!parsed.success) return NextResponse.json({ error: "Invalid booking request" }, { status: 400 });
  const input = parsed.data;
  await checkRateLimitAndThrowError({ identifier: "wnc-native-booking", rateLimitingType: "core" });
  if (input.action === "status") return NextResponse.json(await snapshot(input.requestId));
  if (input.action === "confirmation") {
    const receipt = await snapshot(input.requestId);
    if (
      !receipt ||
      receipt.status !== "confirmed" ||
      receipt.calendarStatus !== "created" ||
      !receipt.meetingUrl
    ) {
      return NextResponse.json({ error: "Microsoft meeting is not ready" }, { status: 409 });
    }
    await submitWncConfirmation({
      ...receipt,
      timeZone: receipt.attendee.timeZone,
      meetingUrl: receipt.meetingUrl,
    });
    return NextResponse.json(await snapshot(input.requestId));
  }

  const expectedDuration =
    input.eventTypeId === Number(configured("WNC_DISCOVERY_EVENT_TYPE_ID") ?? 22)
      ? 15
      : input.eventTypeId === Number(configured("WNC_STRATEGY_EVENT_TYPE_ID") ?? 24)
        ? 60
        : null;
  const eventType = await prisma.eventType.findUnique({
    where: { id: input.eventTypeId },
    select: {
      id: true,
      slug: true,
      length: true,
      requiresConfirmation: true,
      locations: true,
      title: true,
      eventName: true,
      schedulingType: true,
      team: { select: { slug: true } },
      assignAllTeamMembers: true,
      destinationCalendar: {
        select: {
          integration: true,
          credentialId: true,
          credential: { select: { userId: true, type: true, invalid: true } },
        },
      },
      hosts: {
        select: {
          userId: true,
          isFixed: true,
          user: {
            select: {
              name: true,
              email: true,
              destinationCalendar: {
                select: {
                  integration: true,
                  credentialId: true,
                  credential: { select: { userId: true, type: true, invalid: true } },
                },
              },
            },
          },
        },
      },
    },
  });
  const locations = z.array(z.object({ type: z.string() })).safeParse(eventType?.locations);
  if (
    !expectedDuration ||
    !eventType ||
    eventType.team?.slug !== "wealth-navigator" ||
    eventType.length !== expectedDuration ||
    input.durationMinutes !== expectedDuration ||
    eventType.requiresConfirmation ||
    !locations.success ||
    !locations.data.some((location) => location.type === MSTeamsLocationType)
  ) {
    return NextResponse.json({ error: "WNC event configuration needs review" }, { status: 422 });
  }
  const destinations = eventType.hosts.map((host) => ({
    destination: host.user.destinationCalendar,
    userId: host.userId,
  }));
  if (eventType.destinationCalendar)
    destinations.push({ destination: eventType.destinationCalendar, userId: 23 });
  if (
    eventType.assignAllTeamMembers ||
    eventType.hosts.length === 0 ||
    destinations.some(
      ({ destination, userId }) =>
        !destination ||
        destination.integration !== "office365_calendar" ||
        !destination.credentialId ||
        destination.credential?.userId !== userId ||
        destination.credential.type !== "office365_calendar" ||
        destination.credential.invalid === true
    )
  ) {
    return NextResponse.json(
      { error: "Every sales host needs a linked Microsoft booking calendar" },
      { status: 422 }
    );
  }
  const hostIds = eventType.hosts.map((host) => host.userId).sort((a, b) => a - b);
  const expectedHostIds = expectedDuration === 15 ? [23] : [23, 42];
  const approvedName = configured(
    expectedDuration === 15 ? "WNC_FIT_EVENT_NAME" : "WNC_DISCOVERY_EVENT_NAME"
  );
  if (
    JSON.stringify(hostIds) !== JSON.stringify(expectedHostIds) ||
    (expectedDuration === 60 &&
      (eventType.schedulingType !== "COLLECTIVE" || eventType.hosts.some((host) => !host.isFixed))) ||
    !approvedName ||
    eventType.eventName !== approvedName ||
    /[{}]/.test(approvedName.replaceAll("{ATTENDEE}", "")) ||
    !wncConfirmationConfigured()
  ) {
    return NextResponse.json(
      { error: "Founder routing, approved title and Microsoft confirmation setup are required" },
      { status: 422 }
    );
  }
  const organizer = eventType.hosts.find((host) => host.userId === 23)?.user;
  if (organizer?.email !== "quentin@getwealthnavigator.com")
    return NextResponse.json({ error: "Unexpected organizer" }, { status: 422 });
  if (input.action === "details")
    return NextResponse.json({
      titleTemplate: approvedName,
      organizer: { name: organizer.name ?? organizer.email, email: organizer.email },
      participants: eventType.hosts
        .filter((host) => host.userId !== 23)
        .map((host) => ({ name: host.user.name ?? host.user.email, email: host.user.email })),
    });
  if (input.action === "slots") {
    const range = Date.parse(input.end) - Date.parse(input.start);
    if (range <= 0 || range > 8 * 86400_000)
      return NextResponse.json({ error: "Invalid date range" }, { status: 400 });
    const available = await getAvailableSlotsService().getAvailableSlots({
      input: {
        eventTypeId: eventType.id,
        startTime: input.start,
        endTime: input.end,
        timeZone: input.timeZone,
        isTeamEvent: true,
      },
    });
    const slots = Object.values(slotSchema.parse(available.slots))
      .flat()
      .map((slot) => ({
        start: new Date(slot.time).toISOString(),
        end: new Date(Date.parse(slot.time) + expectedDuration * 60_000).toISOString(),
      }));
    return NextResponse.json({ slots });
  }
  const existing = await snapshot(input.requestId);
  if (existing) {
    if (
      existing.eventTypeId !== input.eventTypeId ||
      existing.start !== input.start ||
      existing.attendee.email.toLowerCase() !== input.attendee.email.toLowerCase() ||
      existing.setterEmail !== input.setterEmail ||
      existing.attendee.name !== input.attendee.name ||
      existing.attendee.timeZone !== input.attendee.timeZone
    ) {
      return NextResponse.json({ error: "Retry does not match original booking" }, { status: 409 });
    }
    return NextResponse.json(existing);
  }
  try {
    const bookingData = {
      eventTypeId: input.eventTypeId,
      eventTypeSlug: eventType.slug,
      start: input.start,
      end: new Date(Date.parse(input.start) + expectedDuration * 60_000).toISOString(),
      timeZone: input.attendee.timeZone,
      teamMemberEmail: organizer.email,
      skipContactOwner: true,
      language: "en",
      creationSource: CreationSource.API_V1,
      noEmail: true,
      metadata: {
        wncRequestId: input.requestId,
        wncAttendeeEmail: input.attendee.email,
        wncAttendeeTimeZone: input.attendee.timeZone,
        wncSetterEmail: input.setterEmail,
      },
      responses: {
        name: input.attendee.name,
        email: input.attendee.email,
        location: { value: MSTeamsLocationType, optionValue: "" },
      },
    };
    await getRegularBookingService().createBooking({
      bookingData,
      bookingMeta: {
        userId: -1,
        hostname: req.nextUrl.host,
        idempotencyKey: `wnc:${input.requestId}`,
      },
    });
  } catch {
    // Request metadata is saved before calendar calls. Recover the row after a timeout or racing retry.
    const recovered = await snapshot(input.requestId);
    if (recovered) return NextResponse.json(recovered);
    return NextResponse.json({ error: "Booking failed; reload available times" }, { status: 409 });
  }
  const receipt = await snapshot(input.requestId);
  if (!receipt)
    return NextResponse.json({ error: "Booking receipt unavailable; check booking status" }, { status: 503 });
  return NextResponse.json(receipt);
}
export const POST = defaultResponderForAppDir(handler);
