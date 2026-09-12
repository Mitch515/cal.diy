import { createHash } from "node:crypto";
import { ErrorCode } from "@calcom/lib/errorCodes";
import { ErrorWithCode } from "@calcom/lib/errors";

export interface CalendarReference {
  type: string;
  uid: string;
  credentialId: number | null;
  externalCalendarId: string | null;
}
export interface EvidenceBooking {
  uid: string;
  eventTypeId: number | null;
  status: string;
  startTime: Date;
  endTime: Date;
  references: CalendarReference[];
}

export async function bookingEvidence(
  booking: EvidenceBooking,
  read: (reference: CalendarReference) => Promise<"active" | "absent">
) {
  if (!booking.eventTypeId || booking.eventTypeId < 1)
    throw new ErrorWithCode(ErrorCode.BadRequest, "The original event type is missing");
  const references = booking.references.filter((reference) => reference.type.endsWith("_calendar"));
  if (!references.length)
    throw new ErrorWithCode(ErrorCode.InternalServerError, "Calendar references are missing");
  if (booking.status !== "ACCEPTED" && booking.status !== "CANCELLED") {
    throw new ErrorWithCode(ErrorCode.BadRequest, "Booking status needs separate review");
  }
  const states: ("active" | "absent")[] = [];
  for (const reference of references) states.push(await read(reference));
  const calendarState = states[0];
  if (states.some((state) => state !== calendarState)) {
    throw new ErrorWithCode(ErrorCode.InternalServerError, "Provider retirement is incomplete");
  }
  const referenceFingerprint = createHash("sha256")
    .update(
      JSON.stringify(
        references
          .map((reference) => [
            reference.type,
            reference.uid,
            reference.credentialId,
            reference.externalCalendarId,
          ])
          .sort()
      )
    )
    .digest("hex");
  return {
    uid: booking.uid,
    eventTypeId: booking.eventTypeId,
    referenceFingerprint,
    status: booking.status === "CANCELLED" ? "cancelled" : "scheduled",
    start: booking.startTime.toISOString(),
    end: booking.endTime.toISOString(),
    calendarState,
  };
}
