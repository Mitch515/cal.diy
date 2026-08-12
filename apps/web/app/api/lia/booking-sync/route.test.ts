import { NextRequest } from "next/server";
import { afterEach, beforeEach, describe, expect, test, vi } from "vitest";

const prisma = vi.hoisted(() => ({
  eventType: { findMany: vi.fn() },
  webhook: { upsert: vi.fn(), updateMany: vi.fn() },
  booking: { findMany: vi.fn() },
}));

vi.mock("@calcom/prisma", () => ({ default: prisma }));

import { POST } from "./route";

const APPROVED_URL = "https://sales.example.test/api/webhooks/caldiy";

describe("LIA booking sync", () => {
  beforeEach(() => {
    vi.stubEnv("LIA_INTERNAL_SECRET", "internal-secret");
    vi.stubEnv("LIA_BOOKING_WEBHOOK_URL", APPROVED_URL);
    vi.stubEnv("LIA_BOOKING_ORGANIZER_DOMAIN", "getwealthnavigator.com");
    prisma.eventType.findMany.mockResolvedValue([
      { id: 23, metadata: { clientSlug: "self", clientConfigId: "wn-client" } },
      { id: 17, metadata: null },
    ]);
    prisma.webhook.upsert.mockResolvedValue({ id: "lia-wn-sales-event-23" });
    prisma.webhook.updateMany.mockResolvedValue({ count: 0 });
    prisma.booking.findMany.mockResolvedValue([]);
  });

  afterEach(() => {
    vi.clearAllMocks();
    vi.unstubAllEnvs();
  });

  test("provisions and replays only explicitly WN-sales event types", async () => {
    const response = await POST(request(APPROVED_URL), { params: Promise.resolve({}) });

    expect(response.status).toBe(200);
    expect(await response.json()).toMatchObject({ ok: true, webhookIds: ["lia-wn-sales-event-23"] });
    expect(prisma.eventType.findMany).toHaveBeenCalledWith(
      expect.objectContaining({
        where: {
          owner: { email: { endsWith: "@getwealthnavigator.com", mode: "insensitive" } },
        },
      })
    );
    expect(prisma.webhook.upsert).toHaveBeenCalledWith(
      expect.objectContaining({
        create: expect.objectContaining({ eventTypeId: 23, platform: false, subscriberUrl: APPROVED_URL }),
      })
    );
    expect(prisma.webhook.upsert).toHaveBeenCalledTimes(1);
    expect(prisma.webhook.updateMany).toHaveBeenCalledWith(
      expect.objectContaining({
        where: expect.objectContaining({
          OR: expect.arrayContaining([{ id: { startsWith: "lia-wn-sales-user-" } }]),
        }),
      })
    );
    expect(prisma.booking.findMany).toHaveBeenCalledWith(
      expect.objectContaining({
        where: expect.objectContaining({
          eventTypeId: { in: [23] },
        }),
      })
    );
  });

  test("includes exact event-type provenance in replay snapshots", async () => {
    prisma.booking.findMany.mockResolvedValue([
      {
        id: 37,
        uid: "sales-booking-uid",
        createdAt: new Date("2026-08-11T17:47:50.048Z"),
        updatedAt: new Date("2026-08-12T03:11:30.020Z"),
        startTime: new Date("2026-08-13T17:30:00.000Z"),
        endTime: new Date("2026-08-13T18:30:00.000Z"),
        status: "ACCEPTED",
        location: null,
        responses: null,
        metadata: null,
        fromReschedule: null,
        userPrimaryEmail: "closer@getwealthnavigator.com",
        user: {
          name: "Closer",
          email: "closer@getwealthnavigator.com",
          timeZone: "America/Toronto",
          metadata: null,
        },
        attendees: [
          {
            name: "Advisor",
            email: "advisor@example.test",
            timeZone: "America/Toronto",
            phoneNumber: null,
          },
        ],
        references: [],
        eventType: {
          id: 23,
          slug: "discovery",
          title: "Discovery Call",
          metadata: { clientSlug: "self", clientConfigId: "wn-client" },
        },
      },
    ]);

    const response = await POST(request(APPROVED_URL), { params: Promise.resolve({}) });
    const body = await response.json();

    expect(body.events[0].payload).toMatchObject({
      eventType: { id: 23, slug: "discovery", title: "Discovery Call" },
      metadata: { clientSlug: "self", clientConfigId: "wn-client" },
    });
  });

  test("rejects subscriber takeover before reading any booking data", async () => {
    const response = await POST(request("https://attacker.example.test/webhook"), {
      params: Promise.resolve({}),
    });

    expect(response.status).toBe(403);
    expect(prisma.eventType.findMany).not.toHaveBeenCalled();
    expect(prisma.booking.findMany).not.toHaveBeenCalled();
  });
});

function request(subscriberUrl: string): NextRequest {
  return new NextRequest("https://calendar.example.test/api/lia/booking-sync", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-lia-internal-secret": "internal-secret" },
    body: JSON.stringify({
      subscriberUrl,
      webhookSecret: "signed-webhook-secret",
      updatedAfter: "2026-08-01T00:00:00.000Z",
      limit: 100,
    }),
  });
}
