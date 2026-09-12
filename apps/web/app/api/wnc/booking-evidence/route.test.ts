// @vitest-environment node
import { NextRequest } from "next/server";
import { afterEach, beforeEach, expect, test, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  booking: vi.fn(),
  credential: vi.fn(),
  calendar: vi.fn(),
  event: vi.fn(),
}));
vi.mock("@calcom/prisma", () => ({
  default: { booking: { findUnique: mocks.booking }, credential: { findUnique: mocks.credential } },
}));
vi.mock("@calcom/app-store/googlecalendar/lib/CalendarService", () => ({
  createGoogleCalendarServiceWithGoogleType: () => ({
    authedCalendar: async () => ({ calendars: { get: mocks.calendar }, events: { get: mocks.event } }),
  }),
}));

import { bookingEvidence } from "./evidence";
import { GET } from "./route";

const booking = {
  uid: "fictional-legacy",
  eventTypeId: 22,
  status: "ACCEPTED",
  startTime: new Date("2026-09-14T20:30:00Z"),
  endTime: new Date("2026-09-14T20:45:00Z"),
  eventType: { id: 22, length: 15, team: { slug: "wealth-navigator" } },
  references: [
    {
      type: "google_calendar",
      uid: "google-event",
      credentialId: 1,
      externalCalendarId: "calendar@scenario.invalid",
    },
  ],
};
const request = (secret = "fictional-internal-secret") =>
  new NextRequest("https://cal.example.org/api/wnc/booking-evidence?uid=fictional-legacy", {
    headers: { "x-lia-internal-secret": secret },
  });
beforeEach(() => {
  vi.stubEnv("LIA_INTERNAL_SECRET", "fictional-internal-secret");
  mocks.booking.mockResolvedValue(structuredClone(booking));
  mocks.credential.mockResolvedValue({
    id: 1,
    type: "google_calendar",
    invalid: false,
    delegationCredentialId: null,
  });
  mocks.calendar.mockResolvedValue({ data: { id: "calendar@scenario.invalid" } });
  mocks.event.mockResolvedValue({
    data: {
      id: "google-event",
      status: "confirmed",
      start: { dateTime: booking.startTime.toISOString() },
      end: { dateTime: booking.endTime.toISOString() },
    },
  });
});
afterEach(() => {
  vi.restoreAllMocks();
  vi.resetAllMocks();
  vi.unstubAllEnvs();
});

test("authenticated evidence uses exact provider reads without exposing credentials", async () => {
  const response = await GET(request());
  expect(response.status).toBe(200);
  expect(await response.json()).toMatchObject({
    uid: booking.uid,
    eventTypeId: 22,
    status: "scheduled",
    calendarState: "active",
  });
  expect(mocks.event).toHaveBeenCalledWith({
    calendarId: "calendar@scenario.invalid",
    eventId: "google-event",
  });
});
test("authentication and sales ownership fail before calendar access", async () => {
  expect((await GET(request("wrong"))).status).toBe(403);
  expect(mocks.booking).not.toHaveBeenCalled();
  mocks.booking.mockResolvedValue({
    ...booking,
    eventType: { ...booking.eventType, team: { slug: "other" } },
  });
  expect((await GET(request())).status).toBe(404);
  expect(mocks.event).not.toHaveBeenCalled();
});
test("event absence only counts after the exact calendar remains accessible", async () => {
  mocks.booking.mockResolvedValue({ ...booking, status: "CANCELLED" });
  mocks.event.mockRejectedValue(Object.assign(new Error("Not found"), { code: 404 }));
  expect(await (await GET(request())).json()).toMatchObject({ status: "cancelled", calendarState: "absent" });
  mocks.calendar.mockRejectedValue(Object.assign(new Error("Access removed"), { code: 404 }));
  expect((await GET(request())).status).toBe(503);
});

test("only configured sales event IDs and durations can supply evidence", async () => {
  for (const eventType of [
    { ...booking.eventType, id: 99 },
    { ...booking.eventType, length: 60 },
  ]) {
    mocks.booking.mockResolvedValue({ ...booking, eventType });
    expect((await GET(request())).status).toBe(404);
  }
  vi.stubEnv("WNC_DISCOVERY_EVENT_TYPE_ID", "91");
  mocks.booking.mockResolvedValue({
    ...booking,
    eventTypeId: 91,
    eventType: { ...booking.eventType, id: 91 },
  });
  expect(await (await GET(request())).json()).toMatchObject({ eventTypeId: 91 });
});
test("a cancelled Cal row with an active Google event remains active provider evidence", async () => {
  mocks.booking.mockResolvedValue({ ...booking, status: "CANCELLED" });
  expect(await (await GET(request())).json()).toMatchObject({ status: "cancelled", calendarState: "active" });
});
test("empty, unsupported and incomplete references fail closed", async () => {
  for (const references of [
    [],
    [{ ...booking.references[0], externalCalendarId: null }],
    [{ ...booking.references[0], type: "office365_calendar" }],
  ]) {
    mocks.booking.mockResolvedValue({ ...booking, references });
    expect((await GET(request())).status).toBe(503);
  }
});

test("missing event type cannot authorize repairing an older import", async () => {
  mocks.booking.mockResolvedValue({ ...booking, eventTypeId: null });
  expect((await GET(request())).status).toBe(503);
  expect(mocks.event).not.toHaveBeenCalled();
});
test("reference fingerprints survive ordering and identify removed references", async () => {
  const second = { ...booking.references[0], uid: "second-event" };
  const two = { ...booking, references: [...booking.references, second] };
  const first = await bookingEvidence(two, async () => "absent");
  const reordered = await bookingEvidence(
    { ...two, references: [...two.references].reverse() },
    async () => "absent"
  );
  const removed = await bookingEvidence(booking, async () => "absent");
  expect(first.referenceFingerprint).toBe(reordered.referenceFingerprint);
  expect(first.referenceFingerprint).not.toBe(removed.referenceFingerprint);
  await expect(
    bookingEvidence(two, async (reference) => (reference.uid === "second-event" ? "active" : "absent"))
  ).rejects.toThrow("incomplete");
});
