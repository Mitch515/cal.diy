import { MSTeamsLocationType } from "@calcom/app-store/constants";
import type { CalendarServiceEvent } from "@calcom/types/Calendar";
import type { CredentialForCalendarServiceWithTenantId } from "@calcom/types/Credential";
import { getFixedT } from "i18next";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { z } from "zod";
import BuildCalendarService from "./CalendarService";

const { requestRaw } = vi.hoisted(() => ({
  requestRaw: vi.fn<(args: { url: string; options: RequestInit }) => Promise<Response>>(),
}));

vi.mock("../../_utils/oauth/OAuthManager", () => ({
  OAuthManager: class {
    requestRaw = requestRaw;
  },
}));
vi.mock("../../_utils/oauth/oAuthManagerHelper", () => ({ oAuthManagerHelper: {} }));
vi.mock("@calcom/lib/CalEventParser", () => ({
  getRichDescriptionHTML: () => "<p>Interview</p>",
  getLocation: () => "Microsoft Teams",
}));

const credential = {
  id: 25,
  appId: "office365-calendar",
  type: "office365_calendar",
  userId: 71,
  user: { email: "host@example.com" },
  teamId: null,
  invalid: false,
  delegationCredentialId: null,
  encryptedKey: null,
  key: { access_token: "test", refresh_token: "test", expiry_date: 0, token_type: "Bearer" },
} satisfies CredentialForCalendarServiceWithTenantId;

function booking(): CalendarServiceEvent {
  const language = { translate: getFixedT("en"), locale: "en" };
  return {
    type: "Interview",
    title: "Sales Representative Interview",
    startTime: "2026-09-16T18:45:00Z",
    endTime: "2026-09-16T19:00:00Z",
    calendarDescription: "Interview",
    location: MSTeamsLocationType,
    organizer: { id: 71, name: "Host", email: "host@example.com", timeZone: "America/Toronto", language },
    attendees: [{ name: "Candidate", email: "candidate@example.com", timeZone: "America/Halifax", language }],
    destinationCalendar: [
      {
        id: 17,
        integration: "office365_calendar",
        externalId: "shared/calendar=",
        credentialId: 25,
        userId: null,
        eventTypeId: 32,
        createdAt: null,
        updatedAt: null,
        delegationCredentialId: null,
        primaryEmail: null,
        customCalendarReminder: null,
      },
    ],
  };
}

const payloadSchema = z.object({
  subject: z.string(),
  body: z.object({ content: z.string() }),
  attendees: z.array(z.object({ emailAddress: z.object({ address: z.string() }), type: z.string() })),
  start: z.object({ dateTime: z.string(), timeZone: z.string() }),
  end: z.object({ dateTime: z.string(), timeZone: z.string() }),
  organizer: z.never().optional(),
});

function mutationPayload() {
  const mutation = requestRaw.mock.calls.find(([request]) =>
    ["POST", "PATCH"].includes(request.options.method ?? "")
  );
  if (!mutation || typeof mutation[0].options.body !== "string") throw new Error("Missing event mutation");
  return payloadSchema.parse(JSON.parse(mutation[0].options.body));
}

function ownerResponse(address: string) {
  return new Response(JSON.stringify({ owner: { address } }));
}

describe("Microsoft shared calendar host attendance", () => {
  beforeEach(() => requestRaw.mockReset());

  it("includes the host and candidate when another person owns the destination calendar", async () => {
    requestRaw.mockResolvedValueOnce(ownerResponse("owner@example.com"));
    requestRaw.mockResolvedValueOnce(new Response(JSON.stringify({ id: "event", iCalUId: "ical" })));
    await BuildCalendarService(credential).createEvent(booking(), 25);
    expect(requestRaw.mock.calls[0][0].url).toBe(
      "https://graph.microsoft.com/v1.0/me/calendars/shared%2Fcalendar%3D?$select=owner"
    );
    expect(requestRaw.mock.calls[1][0].url).toBe(
      "https://graph.microsoft.com/v1.0/me/calendars/shared%2Fcalendar%3D/events"
    );
    expect(mutationPayload().attendees).toEqual([
      { emailAddress: { address: "candidate@example.com" }, type: "required" },
      { emailAddress: { address: "host@example.com" }, type: "required" },
    ]);
    expect(mutationPayload().start).toEqual({ dateTime: "2026-09-16T14:45:00", timeZone: "America/Toronto" });
  });

  it("does not invite the host again when they own the calendar", async () => {
    requestRaw.mockResolvedValueOnce(ownerResponse("HOST@example.com"));
    requestRaw.mockResolvedValueOnce(new Response(JSON.stringify({ id: "event" })));
    const event = booking();
    event.destinationCalendar = null;
    await BuildCalendarService(credential).createEvent(event, 25);
    expect(requestRaw.mock.calls[0][0].url).toBe(
      "https://graph.microsoft.com/v1.0/me/calendar?$select=owner"
    );
    expect(mutationPayload().attendees).toHaveLength(1);
  });

  it("does not duplicate an existing host attendee, ignoring email case", async () => {
    requestRaw.mockResolvedValueOnce(ownerResponse("owner@example.com"));
    requestRaw.mockResolvedValueOnce(new Response(JSON.stringify({ id: "event" })));
    const event = booking();
    event.attendees.push({ ...event.organizer, email: "HOST@example.com" });
    await BuildCalendarService(credential).createEvent(event, 25);
    expect(mutationPayload().attendees).toHaveLength(2);
  });

  it("does not create an event when the owner lookup fails", async () => {
    requestRaw.mockResolvedValueOnce(
      new Response(JSON.stringify({ error: { message: "Unavailable" } }), { status: 503 })
    );
    await expect(BuildCalendarService(credential).createEvent(booking(), 25)).rejects.toThrow();
    expect(requestRaw).toHaveBeenCalledTimes(1);
  });

  it("does not silently omit the host when Microsoft returns no owner", async () => {
    requestRaw.mockResolvedValueOnce(new Response("{}"));
    await expect(BuildCalendarService(credential).createEvent(booking(), 25)).rejects.toThrow(
      "owner could not be verified"
    );
    expect(requestRaw).toHaveBeenCalledTimes(1);
  });

  it.each([
    MSTeamsLocationType,
    "Phone call",
  ])("retains host attendance on a %s reschedule", async (location) => {
    requestRaw.mockResolvedValueOnce(
      new Response(
        JSON.stringify({
          organizer: { emailAddress: { address: "owner@example.com" } },
          isOnlineMeeting: true,
          body: { contentType: "html", content: "<a>Existing Teams link</a>" },
        })
      )
    );
    requestRaw.mockResolvedValueOnce(new Response(JSON.stringify({ id: "event" })));
    await BuildCalendarService(credential).updateEvent(
      "event/id=",
      { ...booking(), location },
      "original/shared="
    );
    expect(requestRaw.mock.calls.map(([request]) => request.url)).toEqual([
      "https://graph.microsoft.com/v1.0/me/calendars/original%2Fshared%3D/events/event%2Fid%3D",
      "https://graph.microsoft.com/v1.0/me/calendars/original%2Fshared%3D/events/event%2Fid%3D",
    ]);
    expect(mutationPayload().attendees.map((a) => a.emailAddress.address)).toEqual([
      "candidate@example.com",
      "host@example.com",
    ]);
    if (location === MSTeamsLocationType)
      expect(mutationPayload().body.content).toContain("Existing Teams link");
  });

  it("cancels in the booking's original shared calendar, even if the current destination changed", async () => {
    requestRaw.mockResolvedValueOnce(new Response(null, { status: 204 }));
    await BuildCalendarService(credential).deleteEvent("event/id=", booking(), "original/shared=");
    expect(requestRaw).toHaveBeenCalledWith({
      url: "https://graph.microsoft.com/v1.0/me/calendars/original%2Fshared%3D/events/event%2Fid%3D",
      options: { method: "DELETE" },
    });
  });

  it("supports legacy references without a saved calendar ID", async () => {
    requestRaw.mockResolvedValueOnce(new Response(null, { status: 204 }));
    await BuildCalendarService(credential).deleteEvent("event/id=", booking());
    expect(requestRaw.mock.calls[0][0].url).toBe("https://graph.microsoft.com/v1.0/me/events/event%2Fid%3D");
  });
});
