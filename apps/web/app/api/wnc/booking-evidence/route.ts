import { timingSafeEqual } from "node:crypto";
import process from "node:process";
import { createGoogleCalendarServiceWithGoogleType } from "@calcom/app-store/googlecalendar/lib/CalendarService";
import { ErrorCode } from "@calcom/lib/errorCodes";
import { ErrorWithCode } from "@calcom/lib/errors";
import prisma from "@calcom/prisma";
import { credentialForCalendarServiceSelect } from "@calcom/prisma/selects/credential";
import type { NextRequest } from "next/server";
import { NextResponse } from "next/server";
import { z } from "zod";
import { bookingEvidence, type CalendarReference, type EvidenceBooking } from "./evidence";

const uidSchema = z.string().regex(/^[A-Za-z0-9_-]{8,128}$/);
const metadataSchema = z.object({ clientSlug: z.literal("self"), clientConfigId: z.string().min(1) });

function authorized(supplied: string | null) {
  const configured = Reflect.get(process.env, "LIA_INTERNAL_SECRET");
  if (typeof configured !== "string" || !configured || !supplied) return false;
  const expected = Buffer.from(configured);
  const actual = Buffer.from(supplied);
  return expected.length === actual.length && timingSafeEqual(expected, actual);
}

async function referenceCalendar(reference: CalendarReference) {
  if (reference.type !== "google_calendar" || !reference.externalCalendarId || !reference.credentialId) {
    throw new ErrorWithCode(
      ErrorCode.BadRequest,
      "This legacy reference needs a supported exact calendar reader"
    );
  }
  const credential = await prisma.credential.findUnique({
    where: { id: reference.credentialId },
    select: credentialForCalendarServiceSelect,
  });
  if (
    !credential ||
    credential.invalid ||
    credential.type !== reference.type ||
    credential.delegationCredentialId
  ) {
    throw new ErrorWithCode(ErrorCode.InternalServerError, "The original calendar credential is unavailable");
  }
  const calendar = await createGoogleCalendarServiceWithGoogleType(credential).authedCalendar();
  // An inaccessible calendar's event 404 is not evidence that an invitation was removed.
  const ownedCalendar = await calendar.calendars.get({ calendarId: reference.externalCalendarId });
  if (ownedCalendar.data.id !== reference.externalCalendarId) {
    throw new ErrorWithCode(ErrorCode.InternalServerError, "Legacy calendar identity differs");
  }
  return { calendar, calendarId: reference.externalCalendarId };
}

async function readGoogleReference(
  reference: CalendarReference,
  booking: EvidenceBooking
): Promise<"active" | "absent"> {
  const { calendar, calendarId } = await referenceCalendar(reference);
  try {
    const event = await calendar.events.get({
      calendarId,
      eventId: reference.uid,
    });
    if (event.data.id !== reference.uid)
      throw new ErrorWithCode(ErrorCode.InternalServerError, "Legacy event identity differs");
    if (event.data.status === "cancelled") return "absent";
    assertLegacyTime(event.data.start?.dateTime, event.data.end?.dateTime, booking);
    return "active";
  } catch (error) {
    if (error instanceof Error && "code" in error && (error.code === 404 || error.code === 410))
      return "absent";
    throw error;
  }
}

function assertLegacyTime(
  start: string | null | undefined,
  end: string | null | undefined,
  booking: EvidenceBooking
) {
  if (
    Date.parse(start ?? "") !== booking.startTime.getTime() ||
    Date.parse(end ?? "") !== booking.endTime.getTime()
  ) {
    throw new ErrorWithCode(
      ErrorCode.InternalServerError,
      "Legacy provider time differs from the Cal booking"
    );
  }
}

export async function GET(request: NextRequest) {
  if (!authorized(request.headers.get("x-lia-internal-secret")))
    return NextResponse.json({ error: "Forbidden" }, { status: 403 });
  const parsed = uidSchema.safeParse(request.nextUrl.searchParams.get("uid"));
  if (!parsed.success) return NextResponse.json({ error: "Invalid booking UID" }, { status: 400 });
  try {
    const booking = await prisma.booking.findUnique({
      where: { uid: parsed.data },
      select: {
        uid: true,
        status: true,
        startTime: true,
        endTime: true,
        eventType: { select: { metadata: true } },
        // Deleted flags are deliberately excluded. Only provider readback proves retirement.
        references: { select: { type: true, uid: true, credentialId: true, externalCalendarId: true } },
      },
    });
    if (!booking || !metadataSchema.safeParse(booking.eventType?.metadata).success)
      return NextResponse.json({ error: "Sales booking not found" }, { status: 404 });
    return NextResponse.json(
      await bookingEvidence(booking, (reference) => readGoogleReference(reference, booking)),
      { headers: { "Cache-Control": "no-store" } }
    );
  } catch {
    // Provider exceptions can contain credential-bearing request diagnostics.
    return NextResponse.json({ error: "Legacy calendar evidence is unavailable" }, { status: 503 });
  }
}
