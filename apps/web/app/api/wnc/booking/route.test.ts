// @vitest-environment node
import { NextRequest } from "next/server";
import { afterEach, beforeEach, expect, test, vi } from "vitest";

const mocks = vi.hoisted(() => ({ find: vi.fn(), event: vi.fn(), create: vi.fn(), slots: vi.fn() }));
vi.mock("@calcom/prisma", () => ({
  default: { booking: { findMany: mocks.find }, eventType: { findUnique: mocks.event } },
}));
vi.mock("@calcom/features/bookings/di/RegularBookingService.container", () => ({
  getRegularBookingService: () => ({ createBooking: mocks.create }),
}));
vi.mock("@calcom/features/di/containers/AvailableSlots", () => ({
  getAvailableSlotsService: () => ({ getAvailableSlots: mocks.slots }),
}));
vi.mock("@calcom/lib/checkRateLimitAndThrowError", () => ({ checkRateLimitAndThrowError: vi.fn() }));

import { MSTeamsLocationType } from "@calcom/app-store/constants";
import { POST } from "./route";

const requestId = "a".repeat(64);
const input = {
  action: "book",
  requestId,
  eventTypeId: 22,
  durationMinutes: 15,
  start: "2026-09-14T14:00:00.000Z",
  attendee: { name: "Jamie", email: "jamie@example.org", timeZone: "America/New_York" },
  setterEmail: "bdr@getwealthnavigator.com",
};
const record = {
  uid: "cal-native-1",
  title: "Client acquisition discussion with Jamie",
  eventTypeId: 22,
  startTime: new Date(input.start),
  endTime: new Date("2026-09-14T14:15:00Z"),
  createdAt: new Date("2026-09-10T14:00:00Z"),
  status: "ACCEPTED",
  metadata: {
    wncAttendeeEmail: input.attendee.email,
    wncAttendeeTimeZone: input.attendee.timeZone,
    wncSetterEmail: input.setterEmail,
  },
  user: { name: "Quentin", email: "quentin@getwealthnavigator.com" },
  attendees: [{ name: "Jamie", email: input.attendee.email }],
  references: [
    { type: "office365_calendar", meetingUrl: null },
    { type: "office365_video", meetingUrl: "https://teams.microsoft.com/l/meetup-join/example" },
  ],
};
function call(body: object, secret = "native-internal-secret") {
  return POST(
    new NextRequest("https://cal.example.org/api/wnc/booking", {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-lia-internal-secret": secret },
      body: JSON.stringify(body),
    }),
    { params: Promise.resolve({}) }
  );
}
beforeEach(() => {
  vi.stubEnv("LIA_INTERNAL_SECRET", "native-internal-secret");
  vi.stubEnv("WNC_FIT_EVENT_NAME", "Client acquisition discussion with {ATTENDEE}");
  vi.stubEnv("WNC_DISCOVERY_EVENT_NAME", "Client acquisition discussion with {ATTENDEE}");
  for (const key of ["WNC_MICROSOFT_TENANT_ID", "WNC_MICROSOFT_CLIENT_ID", "WNC_MICROSOFT_CLIENT_SECRET"])
    vi.stubEnv(key, "test-only");
  mocks.event.mockResolvedValue({
    id: 22,
    slug: "advisor-discovery",
    length: 15,
    requiresConfirmation: false,
    locations: [{ type: MSTeamsLocationType }],
    team: { slug: "wealth-navigator" },
    title: "Fit call",
    eventName: "Client acquisition discussion with {ATTENDEE}",
    schedulingType: "ROUND_ROBIN",
    assignAllTeamMembers: false,
    destinationCalendar: null,
    hosts: [
      {
        userId: 23,
        isFixed: false,
        user: {
          name: "Quentin",
          email: "quentin@getwealthnavigator.com",
          destinationCalendar: {
            integration: "office365_calendar",
            credentialId: 18,
            credential: { userId: 23, type: "office365_calendar", invalid: false },
          },
        },
      },
    ],
  });
  mocks.find.mockResolvedValue([]);
  mocks.create.mockResolvedValue({});
});
afterEach(() => {
  vi.resetAllMocks();
  vi.unstubAllEnvs();
});
test("unauthorized calls never query or book", async () => {
  expect((await call(input, "wrong-secret")).status).toBe(403);
  expect(mocks.create).not.toHaveBeenCalled();
  expect(mocks.find).not.toHaveBeenCalled();
});
test("the existing event rules supply native slots", async () => {
  mocks.slots.mockResolvedValue({ slots: { "2026-09-14": [{ time: input.start }] } });
  const response = await call({
    action: "slots",
    eventTypeId: 22,
    durationMinutes: 15,
    start: input.start,
    end: "2026-09-15T14:00:00.000Z",
    timeZone: "America/New_York",
  });
  expect(response.status).toBe(200);
  expect(await response.json()).toEqual({ slots: [{ start: input.start, end: "2026-09-14T14:15:00.000Z" }] });
  expect(mocks.slots).toHaveBeenCalledWith({
    input: expect.objectContaining({ eventTypeId: 22, isTeamEvent: true }),
  });
});
test("uses a server-only durable retry key without internal handoff text", async () => {
  mocks.find.mockResolvedValueOnce([]).mockResolvedValue([record]);
  expect((await call(input)).status).toBe(200);
  expect(mocks.create).toHaveBeenCalledWith(
    expect.objectContaining({
      bookingMeta: expect.objectContaining({ idempotencyKey: `wnc:${requestId}` }),
      bookingData: expect.objectContaining({
        eventTypeId: 22,
        noEmail: true,
        metadata: expect.objectContaining({ wncRequestId: requestId, wncConfirmationEligible: "true" }),
      }),
    })
  );
  expect(JSON.stringify(mocks.create.mock.calls)).not.toContain("handoff");
});

test("reads the request metadata after Cal replaces the slot collision key", async () => {
  mocks.find.mockResolvedValue([
    { ...record, idempotencyKey: "cal-slot-key", metadata: { ...record.metadata, wncRequestId: requestId } },
  ]);
  const response = await call({ action: "status", requestId });
  expect(response.status).toBe(200);
  expect(await response.json()).toMatchObject({ uid: record.uid });
  expect(mocks.find).toHaveBeenCalledWith(
    expect.objectContaining({
      where: {
        OR: [
          { idempotencyKey: `wnc:${requestId}` },
          { metadata: { path: ["wncRequestId"], equals: requestId }, fromReschedule: null },
        ],
      },
      take: 2,
    })
  );
  expect(mocks.create).not.toHaveBeenCalled();
});

test("a cancelled booking remains recoverable after Cal clears the collision key", async () => {
  mocks.find.mockResolvedValue([
    {
      ...record,
      status: "CANCELLED",
      idempotencyKey: null,
      metadata: { ...record.metadata, wncRequestId: requestId },
    },
  ]);
  const response = await call(input);
  expect(await response.json()).toMatchObject({ uid: record.uid, status: "cancelled" });
  expect(mocks.create).not.toHaveBeenCalled();
});

test("ambiguous request identities fail closed without creating another meeting", async () => {
  mocks.find.mockResolvedValue([record, { ...record, uid: "different-booking" }]);
  expect((await call(input)).status).toBe(500);
  expect(mocks.create).not.toHaveBeenCalled();
});

test("a missing post-create receipt never returns a successful null response", async () => {
  const response = await call(input);
  expect(response.status).toBe(503);
  expect(await response.json()).toMatchObject({ error: "Booking receipt unavailable; check booking status" });
  expect(mocks.create).toHaveBeenCalledTimes(1);
});
test("an existing key is recovered without creating another invitation", async () => {
  mocks.find.mockResolvedValue([record]);
  const response = await call(input);
  expect(response.status).toBe(200);
  expect(mocks.create).not.toHaveBeenCalled();
});
test("a different setter cannot reuse the original request", async () => {
  mocks.find.mockResolvedValue([record]);
  expect((await call({ ...input, setterEmail: "another@getwealthnavigator.com" })).status).toBe(409);
  expect(mocks.create).not.toHaveBeenCalled();
});
test("unrelated events and duration drift fail closed", async () => {
  expect((await call({ ...input, eventTypeId: 999 })).status).toBe(422);
  expect((await call({ ...input, durationMinutes: 60 })).status).toBe(422);
  expect(mocks.create).not.toHaveBeenCalled();
});

test("a Google destination is blocked before any invitation", async () => {
  mocks.event.mockResolvedValue({
    id: 22,
    slug: "advisor-discovery",
    length: 15,
    requiresConfirmation: false,
    locations: [{ type: MSTeamsLocationType }],
    team: { slug: "wealth-navigator" },
    assignAllTeamMembers: false,
    destinationCalendar: null,
    hosts: [{ user: { destinationCalendar: { integration: "google_calendar", credentialId: 16 } } }],
  });
  expect((await call(input)).status).toBe(422);
  expect(mocks.create).not.toHaveBeenCalled();
});

test("the receipt retains the title and actual participants without adding the setter", async () => {
  mocks.find.mockResolvedValue([
    {
      ...record,
      attendees: [...record.attendees, { name: "Second closer", email: "second@getwealthnavigator.com" }],
    },
  ]);
  const response = await call({ action: "status", requestId });
  expect(await response.json()).toMatchObject({
    title: record.title,
    participants: [{ name: "Second closer", email: "second@getwealthnavigator.com" }],
    guests: ["second@getwealthnavigator.com"],
    calendarStatus: "created",
  });
});

test.each([
  { userId: 42, type: "office365_calendar", invalid: false },
  { userId: 23, type: "office365_calendar", invalid: true },
  { userId: 23, type: "google_calendar", invalid: false },
])("rejects a mismatched or invalid Microsoft credential before booking: %j", async (credential) => {
  mocks.event.mockResolvedValue({
    id: 22,
    length: 15,
    requiresConfirmation: false,
    locations: [{ type: MSTeamsLocationType }],
    team: { slug: "wealth-navigator" },
    assignAllTeamMembers: false,
    destinationCalendar: null,
    hosts: [
      {
        userId: 23,
        user: { destinationCalendar: { integration: "office365_calendar", credentialId: 18, credential } },
      },
    ],
  });
  expect((await call(input)).status).toBe(422);
  expect(mocks.create).not.toHaveBeenCalled();
});

test("discovery requires both founders collectively and describes both before booking", async () => {
  const founder = (userId: number, name: string, email: string) => ({
    userId,
    isFixed: true,
    user: {
      name,
      email,
      destinationCalendar: {
        integration: "office365_calendar",
        credentialId: userId === 23 ? 18 : 21,
        credential: { userId, type: "office365_calendar", invalid: false },
      },
    },
  });
  const event = {
    id: 24,
    slug: "advisor-strategy",
    length: 60,
    requiresConfirmation: false,
    title: "Discovery call",
    eventName: "Client acquisition discussion with {ATTENDEE}",
    schedulingType: "COLLECTIVE",
    locations: [{ type: MSTeamsLocationType }],
    team: { slug: "wealth-navigator" },
    assignAllTeamMembers: false,
    destinationCalendar: null,
    hosts: [
      founder(23, "Quentin", "quentin@getwealthnavigator.com"),
      founder(42, "Louay", "louay@getwealthnavigator.com"),
    ],
  };
  mocks.event.mockResolvedValue(event);
  const details = await call({ action: "details", eventTypeId: 24, durationMinutes: 60 });
  expect(details.status).toBe(200);
  expect(await details.json()).toMatchObject({
    organizer: { email: "quentin@getwealthnavigator.com" },
    participants: [{ email: "louay@getwealthnavigator.com" }],
  });
  mocks.event.mockResolvedValue({ ...event, schedulingType: "ROUND_ROBIN" });
  expect((await call({ action: "details", eventTypeId: 24, durationMinutes: 60 })).status).toBe(422);
  mocks.event.mockResolvedValue({ ...event, hosts: [event.hosts[0]] });
  expect((await call({ action: "details", eventTypeId: 24, durationMinutes: 60 })).status).toBe(422);
});

test("a confirmation request for a booking made before the release is refused before any claim", async () => {
  mocks.find.mockResolvedValue([
    { ...record, metadata: { ...record.metadata, wncRequestId: requestId, wncConfirmation: { key: "k", state: "preparing" } } },
  ]);
  const response = await call({ action: "confirmation", requestId });
  expect(response.status).toBe(400);
  expect(await response.text()).toContain("predates automatic confirmations");
});
