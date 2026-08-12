import { NextRequest } from "next/server";
import { afterEach, beforeEach, describe, expect, test, vi } from "vitest";

const prisma = vi.hoisted(() => ({
  user: { findMany: vi.fn() },
  webhook: { findUnique: vi.fn(), upsert: vi.fn(), updateMany: vi.fn() },
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
    prisma.user.findMany.mockResolvedValue([{ id: 25 }]);
    prisma.webhook.findUnique.mockResolvedValue(null);
    prisma.webhook.upsert.mockResolvedValue({ id: "lia-wn-sales-user-25" });
    prisma.webhook.updateMany.mockResolvedValue({ count: 0 });
    prisma.booking.findMany.mockResolvedValue([]);
  });

  afterEach(() => {
    vi.clearAllMocks();
    vi.unstubAllEnvs();
  });

  test("provisions only organizer-scoped webhooks and scopes booking replay to the same domain", async () => {
    const response = await POST(request(APPROVED_URL), { params: Promise.resolve({}) });

    expect(response.status).toBe(200);
    expect(await response.json()).toMatchObject({ ok: true, webhookIds: ["lia-wn-sales-user-25"] });
    expect(prisma.user.findMany).toHaveBeenCalledWith(
      expect.objectContaining({
        where: { email: { endsWith: "@getwealthnavigator.com", mode: "insensitive" } },
      })
    );
    expect(prisma.webhook.upsert).toHaveBeenCalledWith(
      expect.objectContaining({
        create: expect.objectContaining({ userId: 25, platform: false, subscriberUrl: APPROVED_URL }),
      })
    );
    expect(prisma.booking.findMany).toHaveBeenCalledWith(
      expect.objectContaining({
        where: expect.objectContaining({
          user: { email: { endsWith: "@getwealthnavigator.com", mode: "insensitive" } },
        }),
      })
    );
  });

  test("rejects subscriber takeover before reading any booking data", async () => {
    const response = await POST(request("https://attacker.example.test/webhook"), {
      params: Promise.resolve({}),
    });

    expect(response.status).toBe(403);
    expect(prisma.user.findMany).not.toHaveBeenCalled();
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
