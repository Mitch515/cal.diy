import { PrismaApiKeyRepository } from "@calcom/features/api-keys-legacy/api-keys/repositories/PrismaApiKeyRepository";
import { DEFAULT_SCHEDULE, getAvailabilityFromSchedule } from "@calcom/lib/availability";
import { emailRegex } from "@calcom/lib/emailSchema";
import slugify from "@calcom/lib/slugify";
import prisma from "@calcom/prisma";
import type { Prisma } from "@calcom/prisma/client";
import { CreationSource, IdentityProvider } from "@calcom/prisma/enums";
import { defaultResponderForAppDir } from "app/api/defaultResponderForAppDir";
import type { NextRequest } from "next/server";
import { NextResponse } from "next/server";
import z from "zod";

const provisionHostSchema = z.object({
  email: z
    .string()
    .regex(emailRegex)
    .transform((value) => value.toLowerCase()),
  name: z.string().trim().min(1).optional(),
  username: z.string().trim().min(1).optional(),
  timeZone: z.string().trim().min(1).default("America/Toronto"),
  brandColor: z.string().trim().min(1).optional(),
  theme: z.string().trim().min(1).optional(),
  metadata: z.record(z.unknown()).optional(),
  publicBookingBaseUrl: z.string().trim().min(1).optional(),
});

async function handler(req: NextRequest) {
  const configuredSecret = process.env.LIA_INTERNAL_SECRET ?? process.env.NEXTAUTH_SECRET;
  if (!configuredSecret) {
    return NextResponse.json({ message: "LIA provisioning is not configured" }, { status: 503 });
  }

  if (req.headers.get("x-lia-internal-secret") !== configuredSecret) {
    return NextResponse.json({ message: "Invalid LIA provisioning secret" }, { status: 403 });
  }

  const parsed = provisionHostSchema.safeParse(await req.json().catch(() => null));
  if (!parsed.success) {
    return NextResponse.json({ message: parsed.error.message }, { status: 400 });
  }

  const input = parsed.data;
  const user = await upsertProvisionedUser(input);
  const repository = await PrismaApiKeyRepository.withGlobalPrisma();
  const apiKey = await repository.createApiKey({
    userId: user.id,
    note: "LIA automated host provisioning",
    expiresAt: null,
  });

  return NextResponse.json({
    apiKey,
    user: {
      id: user.id,
      username: user.username,
      email: user.email,
      name: user.name,
      timeZone: user.timeZone,
    },
  });
}

export const POST = defaultResponderForAppDir(handler);

async function upsertProvisionedUser(input: z.infer<typeof provisionHostSchema>) {
  const existingUser = await prisma.user.findUnique({
    where: { email: input.email },
    select: {
      id: true,
      username: true,
      email: true,
      name: true,
      timeZone: true,
      defaultScheduleId: true,
      metadata: true,
    },
  });

  if (existingUser) {
    const username =
      existingUser.username ?? (await uniqueUsername(input.username ?? input.name ?? input.email));
    const updatedUser = await prisma.user.update({
      where: { id: existingUser.id },
      data: {
        username,
        name: input.name ?? existingUser.name,
        timeZone: input.timeZone,
        emailVerified: new Date(),
        completedOnboarding: true,
        locale: "en",
        ...(input.brandColor ? { brandColor: input.brandColor } : {}),
        ...(input.theme ? { theme: input.theme } : {}),
        ...(input.metadata !== undefined
          ? { metadata: mergeMetadata(existingUser.metadata, input.metadata) }
          : {}),
      },
      select: provisionedUserSelect,
    });
    await ensureDefaultSchedule(updatedUser.id, updatedUser.defaultScheduleId, input.timeZone);
    return updatedUser;
  }

  const username = await uniqueUsername(input.username ?? input.name ?? input.email);
  const createdUser = await prisma.user.create({
    data: {
      username,
      email: input.email,
      name: input.name ?? null,
      timeZone: input.timeZone,
      emailVerified: new Date(),
      completedOnboarding: true,
      locale: "en",
      identityProvider: IdentityProvider.CAL,
      creationSource: CreationSource.WEBAPP,
      ...(input.brandColor ? { brandColor: input.brandColor } : {}),
      ...(input.theme ? { theme: input.theme } : {}),
      ...(input.metadata !== undefined ? { metadata: input.metadata as Prisma.InputJsonObject } : {}),
      schedules: {
        create: {
          name: "Working Hours",
          timeZone: input.timeZone,
          availability: {
            createMany: {
              data: getAvailabilityFromSchedule(DEFAULT_SCHEDULE),
            },
          },
        },
      },
    },
    select: provisionedUserSelect,
  });
  await ensureDefaultSchedule(createdUser.id, createdUser.defaultScheduleId, input.timeZone);
  return createdUser;
}

const provisionedUserSelect = {
  id: true,
  username: true,
  email: true,
  name: true,
  timeZone: true,
  defaultScheduleId: true,
} as const;

async function ensureDefaultSchedule(userId: number, defaultScheduleId: number | null, timeZone: string) {
  if (defaultScheduleId) return;

  const existingSchedule = await prisma.schedule.findFirst({
    where: { userId },
    select: { id: true },
    orderBy: { id: "asc" },
  });
  const scheduleId =
    existingSchedule?.id ??
    (
      await prisma.schedule.create({
        data: {
          userId,
          name: "Working Hours",
          timeZone,
          availability: {
            createMany: {
              data: getAvailabilityFromSchedule(DEFAULT_SCHEDULE),
            },
          },
        },
        select: { id: true },
      })
    ).id;

  await prisma.user.update({
    where: { id: userId },
    data: { defaultScheduleId: scheduleId },
    select: { id: true },
  });
}

async function uniqueUsername(value: string) {
  const base = slugify(value.includes("@") ? (value.split("@")[0] ?? value) : value);
  const fallback = base || "lia-host";

  for (let attempt = 0; attempt < 20; attempt += 1) {
    const candidate = attempt === 0 ? fallback : `${fallback}-${attempt + 1}`;
    const existing = await prisma.user.findFirst({
      where: { username: candidate },
      select: { id: true },
    });
    if (!existing) return candidate;
  }

  return `${fallback}-${Date.now()}`;
}

function mergeMetadata(existingMetadata: Prisma.JsonValue | null, incomingMetadata: Record<string, unknown>) {
  const existingObject =
    existingMetadata && typeof existingMetadata === "object" && !Array.isArray(existingMetadata)
      ? existingMetadata
      : {};

  return {
    ...existingObject,
    ...incomingMetadata,
  } as Prisma.InputJsonObject;
}
