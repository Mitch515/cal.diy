import { OUTLOOK_TENANT_ID } from "@calcom/features/auth/lib/outlook";
import type { CredentialForCalendarServiceWithTenantId } from "@calcom/types/Credential";
import { describe, expect, it, vi } from "vitest";
import BuildCalendarService from "./CalendarService";

type OAuthManagerOptions = {
  fetchNewTokenObject: (args: { refreshToken: string | null }) => Promise<Response | null>;
};

const { oauthManagerOptions }: { oauthManagerOptions: { current: OAuthManagerOptions | null } } = vi.hoisted(
  () => ({
    oauthManagerOptions: {
      current: null,
    },
  })
);

vi.mock("../../_utils/oauth/OAuthManager", () => ({
  OAuthManager: vi.fn().mockImplementation(function MockOAuthManager(options: OAuthManagerOptions) {
    oauthManagerOptions.current = options;
    return { requestRaw: vi.fn() };
  }),
}));

vi.mock("./getOfficeAppKeys", () => ({
  getOfficeAppKeys: vi.fn().mockResolvedValue({
    client_id: "test-client-id",
    client_secret: "test-client-secret",
  }),
}));

describe("Office365CalendarService token refresh", () => {
  it("uses the configured tenant endpoint for a personal credential", async () => {
    const fetchMock = vi.fn().mockResolvedValue(new Response(null, { status: 200 }));
    vi.stubGlobal("fetch", fetchMock);

    BuildCalendarService({
      appId: "office365-calendar",
      delegatedTo: null,
      delegatedToId: null,
      delegationCredentialId: null,
      encryptedKey: null,
      id: 1,
      invalid: false,
      key: {
        access_token: "expired-access-token",
        expiry_date: 0,
        refresh_token: "refresh-token",
        token_type: "Bearer",
      },
      teamId: null,
      type: "office365_calendar",
      user: { email: "advisor@example.com" },
      userId: 1,
    } as CredentialForCalendarServiceWithTenantId);

    await oauthManagerOptions.current?.fetchNewTokenObject({ refreshToken: "refresh-token" });

    expect(fetchMock).toHaveBeenCalledWith(
      `https://login.microsoftonline.com/${OUTLOOK_TENANT_ID}/oauth2/v2.0/token`,
      expect.objectContaining({ method: "POST" })
    );
    expect(OUTLOOK_TENANT_ID).not.toBe("common");
  });
});
