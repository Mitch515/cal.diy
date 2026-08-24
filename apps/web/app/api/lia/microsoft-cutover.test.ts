import { MSTeamsLocationType } from "@calcom/app-store/constants";
import { MICROSOFT_CALENDAR_AND_TEAMS_SCOPES, WEBAPP_URL } from "@calcom/lib/constants";
import { NextRequest } from "next/server";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  setDefaultConferencingApp: vi.fn(),
  prisma: {
    user: { findFirst: vi.fn(), update: vi.fn() },
    credential: {
      findFirst: vi.fn(),
      findMany: vi.fn(),
      update: vi.fn(),
      create: vi.fn(),
      deleteMany: vi.fn(),
    },
    eventType: { findFirst: vi.fn(), findUnique: vi.fn(), create: vi.fn(), update: vi.fn() },
  },
}));

vi.mock("@calcom/prisma", () => ({ default: mocks.prisma }));
vi.mock("@calcom/app-store/_utils/setDefaultConferencingApp", () => ({
  default: mocks.setDefaultConferencingApp,
}));

import { ensureMicrosoftTeamsConnection } from "@calcom/app-store/office365calendar/lib/ensureMicrosoftTeamsConnection";
import { POST as upsertEventType } from "./event-types/route";
import { resolveConnectTarget } from "./host-session/route";

describe("LIA Microsoft cutover", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mocks.prisma.credential.deleteMany.mockResolvedValue({ count: 0 });
    mocks.setDefaultConferencingApp.mockResolvedValue(undefined);
  });

  afterEach(() => {
    vi.unstubAllEnvs();
  });

  it("sends the host to one consent covering Outlook Calendar and Teams", async () => {
    const target = await resolveConnectTarget(
      31,
      "alexandre@getwealthnavigator.com",
      "session-secret",
      async () => ({ client_id: "microsoft-client-id" })
    );
    const url = new URL(target);

    expect(url.origin).toBe("https://login.microsoftonline.com");
    expect(url.pathname).toBe("/common/oauth2/v2.0/authorize");
    expect(url.searchParams.get("client_id")).toBe("microsoft-client-id");
    expect(url.searchParams.get("scope")?.split(" ")).toEqual(MICROSOFT_CALENDAR_AND_TEAMS_SCOPES);
    expect(url.searchParams.get("login_hint")).toBe("alexandre@getwealthnavigator.com");
    expect(url.searchParams.get("redirect_uri")).toContain("/api/integrations/office365calendar/callback");

    const state = JSON.parse(url.searchParams.get("state") ?? "{}") as {
      returnTo?: string;
      onErrorReturnTo?: string;
      nonce?: string;
      nonceHash?: string;
    };
    expect(state.returnTo).toBe(new URL("/lia/connected", WEBAPP_URL).toString());
    expect(state.onErrorReturnTo).toBe(state.returnTo);
    expect(state.nonce).toBeTruthy();
    expect(state.nonceHash).toMatch(/^[a-f0-9]{64}$/);
  });

  it("falls back to calendar settings when Microsoft is not configured", async () => {
    const target = await resolveConnectTarget(
      31,
      "alexandre@getwealthnavigator.com",
      "session-secret",
      async () => ({})
    );

    expect(target).toBe(new URL("/apps/installed/calendar", WEBAPP_URL).toString());
  });

  it("updates the stable Teams credential and removes stale duplicates", async () => {
    mocks.prisma.credential.findMany.mockResolvedValue([{ id: 12 }, { id: 19 }]);
    mocks.prisma.credential.update.mockResolvedValue({ id: 12 });

    const result = await ensureMicrosoftTeamsConnection({
      userId: 26,
      key: { access_token: "access-token", refresh_token: "refresh-token" },
    });

    expect(result).toEqual({ id: 12 });
    expect(mocks.prisma.credential.update).toHaveBeenCalledWith({
      where: { id: 12 },
      data: {
        key: { access_token: "access-token", refresh_token: "refresh-token" },
        invalid: false,
      },
      select: { id: true },
    });
    expect(mocks.prisma.credential.deleteMany).toHaveBeenCalledWith({
      where: { id: { in: [19] }, userId: 26, type: "office365_video" },
    });
    expect(mocks.setDefaultConferencingApp).toHaveBeenCalledWith(26, "msteams");
  });

  it("creates the Teams credential on a user's first combined Microsoft consent", async () => {
    mocks.prisma.credential.findMany.mockResolvedValue([]);
    mocks.prisma.credential.create.mockResolvedValue({ id: 44 });

    const result = await ensureMicrosoftTeamsConnection({
      userId: 31,
      key: { access_token: "access-token" },
    });

    expect(result).toEqual({ id: 44 });
    expect(mocks.prisma.credential.create).toHaveBeenCalledWith({
      data: {
        userId: 31,
        appId: "msteams",
        type: "office365_video",
        key: { access_token: "access-token" },
      },
      select: { id: true },
    });
    expect(mocks.setDefaultConferencingApp).toHaveBeenCalledWith(31, "msteams");
  });

  it("hard-codes future LIA meetings to the connected host's Teams credential", async () => {
    vi.stubEnv("LIA_INTERNAL_SECRET", "internal-secret");
    mocks.prisma.user.findFirst.mockResolvedValue({ id: 31, username: "alexandre", metadata: null });
    mocks.prisma.credential.findFirst.mockResolvedValue({ id: 44 });
    mocks.prisma.eventType.findUnique.mockResolvedValue(null);
    mocks.prisma.eventType.create.mockResolvedValue({
      id: 73,
      title: "Discovery Call",
      slug: "discovery",
      length: 15,
      owner: { username: "alexandre" },
    });
    mocks.prisma.eventType.update.mockResolvedValue({ id: 73 });

    const response = await upsertEventType(
      new NextRequest("https://calendar.example.test/api/lia/event-types", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-lia-internal-secret": "internal-secret",
        },
        body: JSON.stringify({
          email: "alexandre@getwealthnavigator.com",
          title: "Discovery Call",
          slug: "discovery",
          durationMinutes: 15,
          timeZone: "America/Edmonton",
        }),
      }),
      { params: Promise.resolve({}) }
    );

    expect(response.status).toBe(200);
    expect(mocks.prisma.eventType.create).toHaveBeenCalledWith(
      expect.objectContaining({
        data: expect.objectContaining({
          locations: [{ type: MSTeamsLocationType, credentialId: 44 }],
        }),
      })
    );
  });
});
