import type { CredentialForCalendarServiceWithTenantId } from "@calcom/types/Credential";
import { beforeEach, expect, test, vi } from "vitest";

const { request } = vi.hoisted(() => ({ request: vi.fn() }));
vi.mock("../../_utils/oauth/OAuthManager", () => ({
  OAuthManager: vi.fn().mockImplementation(function MockOAuthManager() {
    return { requestRaw: request };
  }),
}));

import { readOffice365Reference } from "./CalendarService";

const credential = {
  appId: "office365-calendar",
  delegatedTo: null,
  delegatedToId: null,
  delegationCredentialId: null,
  encryptedKey: null,
  id: 1,
  invalid: false,
  key: {
    access_token: "fictional-token",
    expiry_date: 0,
    refresh_token: "fictional-refresh",
    token_type: "Bearer",
  },
  teamId: null,
  type: "office365_calendar",
  user: { email: "advisor@example.com" },
  userId: 1,
} satisfies CredentialForCalendarServiceWithTenantId;
const event = {
  id: "event/id",
  isCancelled: false,
  start: { dateTime: "2026-09-17T19:00:00.0000000", timeZone: "UTC" },
  end: { dateTime: "2026-09-17T19:15:00.0000000", timeZone: "UTC" },
};
beforeEach(() => request.mockReset());
function calendar() {
  request.mockResolvedValueOnce(Response.json({ id: "calendar/id" }));
}
const read = () => readOffice365Reference(credential, "event/id", "calendar/id");

test("reads the exact calendar before the exact event, using UTC and only GET", async () => {
  calendar();
  request.mockResolvedValueOnce(Response.json(event));
  expect(await read()).toEqual({
    state: "active",
    start: "2026-09-17T19:00:00.0000000Z",
    end: "2026-09-17T19:15:00.0000000Z",
  });
  expect(request.mock.calls.map((call) => call[0])).toEqual([
    {
      url: "https://graph.microsoft.com/v1.0/me/calendars/calendar%2Fid?$select=id",
      options: { method: "GET" },
    },
    {
      url: "https://graph.microsoft.com/v1.0/me/calendars/calendar%2Fid/events/event%2Fid",
      options: { method: "GET", headers: { Prefer: 'outlook.timezone="UTC"' } },
    },
  ]);
});
test.each([404, 410])("event %s proves absence after a successful calendar read", async (status) => {
  calendar();
  request.mockResolvedValueOnce(new Response(null, { status }));
  expect(await read()).toEqual({ state: "absent" });
});
test("an inaccessible calendar never proves event absence", async () => {
  request.mockResolvedValueOnce(Response.json({ error: { message: "Unavailable" } }, { status: 404 }));
  await expect(read()).rejects.toThrow();
  expect(request).toHaveBeenCalledTimes(1);
});
test("calendar identity mismatch prevents event access", async () => {
  request.mockResolvedValueOnce(Response.json({ id: "other" }));
  await expect(read()).rejects.toThrow("calendar identity");
  expect(request).toHaveBeenCalledTimes(1);
});
test("event identity, malformed UTC evidence and authorization failures are not absence", async () => {
  for (const body of [
    { ...event, id: "other" },
    { ...event, start: { ...event.start, timeZone: "America/New_York" } },
    { ...event, end: {} },
  ]) {
    calendar();
    request.mockResolvedValueOnce(Response.json(body));
    await expect(read()).rejects.toThrow();
  }
  calendar();
  request.mockResolvedValueOnce(Response.json({ error: { message: "Denied" } }, { status: 403 }));
  await expect(read()).rejects.toThrow();
});
test("a cancelled provider event proves absence", async () => {
  calendar();
  request.mockResolvedValueOnce(Response.json({ ...event, isCancelled: true }));
  expect(await read()).toEqual({ state: "absent" });
});
