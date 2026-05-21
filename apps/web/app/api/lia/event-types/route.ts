import { defaultResponderForAppDir } from "app/api/defaultResponderForAppDir";
import type { NextRequest } from "next/server";
import { NextResponse } from "next/server";
import z from "zod";

import { emailRegex } from "@calcom/lib/emailSchema";
import prisma from "@calcom/prisma";

const upsertEventTypeSchema = z.object({
  id: z.number().int().positive().optional(),
  userId: z.number().int().positive().optional(),
  username: z.string().trim().min(1).optional(),
  email: z
    .string()
    .regex(emailRegex)
    .transform((value) => value.toLowerCase())
    .optional(),
  title: z.string().trim().min(1),
  slug: z.string().trim().min(1),
  description: z.string().trim().nullable().optional(),
  durationMinutes: z.number().int().min(5).max(240),
  timeZone: z.string().trim().min(1),
  status: z.enum(["draft", "active", "disabled"]).default("active"),
}).refine((input) => input.userId || input.username || input.email, {
  message: "userId, username, or email is required",
});

async function handler(req: NextRequest) {
  const configuredSecret =
    process.env.LIA_INTERNAL_SECRET ?? process.env.NEXTAUTH_SECRET;
  if (!configuredSecret) {
    return NextResponse.json(
      { message: "LIA provisioning is not configured" },
      { status: 503 },
    );
  }

  if (req.headers.get("x-lia-internal-secret") !== configuredSecret) {
    return NextResponse.json(
      { message: "Invalid LIA provisioning secret" },
      { status: 403 },
    );
  }

  const parsed = upsertEventTypeSchema.safeParse(
    await req.json().catch(() => null),
  );
  if (!parsed.success) {
    return NextResponse.json(
      { message: parsed.error.message },
      { status: 400 },
    );
  }

  const eventType = await upsertLiaEventType(parsed.data);
  if (!eventType) {
    return NextResponse.json(
      { message: "Cal.diy host was not found" },
      { status: 404 },
    );
  }

  return NextResponse.json({
    id: eventType.id,
    username: eventType.owner?.username ?? null,
    slug: eventType.slug,
    title: eventType.title,
    durationMinutes: eventType.length,
  });
}

export const POST = defaultResponderForAppDir(handler);

async function upsertLiaEventType(
  input: z.infer<typeof upsertEventTypeSchema>,
) {
  const user = await prisma.user.findFirst({
    where: {
      ...(input.userId ? { id: input.userId } : {}),
      ...(input.username ? { username: input.username } : {}),
      ...(input.email ? { email: input.email } : {}),
    },
    select: { id: true, username: true },
  });

  if (!user) {
    return null;
  }

  const existingEventType = input.id
    ? await prisma.eventType.findFirst({
        where: { id: input.id, userId: user.id },
        select: { id: true },
      })
    : await prisma.eventType.findUnique({
        where: { userId_slug: { userId: user.id, slug: input.slug } },
        select: { id: true },
      });

  const hidden = input.status !== "active";
  const eventTypeData = {
    title: input.title,
    slug: input.slug,
    description: input.description ?? null,
    length: input.durationMinutes,
    hidden,
    timeZone: input.timeZone,
  };

  const eventType = existingEventType
    ? await prisma.eventType.update({
        where: { id: existingEventType.id },
        data: eventTypeData,
        select: eventTypeSelect,
      })
    : await prisma.eventType.create({
        data: {
          ...eventTypeData,
          userId: user.id,
          users: { connect: { id: user.id } },
        },
        select: eventTypeSelect,
      });

  await prisma.eventType.update({
    where: { id: eventType.id },
    data: { users: { connect: { id: user.id } } },
    select: { id: true },
  });

  return eventType;
}

const eventTypeSelect = {
  id: true,
  title: true,
  slug: true,
  length: true,
  owner: { select: { username: true } },
} as const;
