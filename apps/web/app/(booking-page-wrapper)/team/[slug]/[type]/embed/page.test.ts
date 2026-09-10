// @vitest-environment node
import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  getServerSideProps: vi.fn(),
  loadTranslations: vi.fn(),
  buildLegacyCtx: vi.fn(),
}));

vi.mock("next/headers", () => ({ headers: async () => ({}), cookies: async () => ({}) }));
vi.mock("next/navigation", () => ({
  notFound: () => {
    throw new Error("NEXT_NOT_FOUND");
  },
  redirect: (url: string) => {
    throw new Error(`NEXT_REDIRECT:${url}`);
  },
}));
vi.mock("app/CustomI18nProvider", () => ({ CustomI18nProvider: () => null }));
vi.mock("@calcom/i18n/server", () => ({ loadTranslations: mocks.loadTranslations }));
vi.mock("@calcom/web/modules/team/type-view", () => ({ default: () => null }));
vi.mock("@lib/buildLegacyCtx", () => ({ buildLegacyCtx: mocks.buildLegacyCtx }));
vi.mock("@lib/team/[slug]/[type]/getServerSideProps", () => ({
  getServerSideProps: mocks.getServerSideProps,
}));

import TeamBookingEmbedPage, { generateMetadata } from "./page";

const params = { slug: "wealth-navigator", type: "advisor-discovery" };
const query = { embed: "wnc-booking", layout: "month_view", email: "advisor@example.invalid" };
const input = { params: Promise.resolve(params), searchParams: Promise.resolve(query) };

describe("team booking embed route", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mocks.buildLegacyCtx.mockReturnValue({ params, query });
    mocks.getServerSideProps.mockResolvedValue({ props: { eventData: { interfaceLanguage: null } } });
  });

  it("renders the team booking view in embed mode with the original team and prefill context", async () => {
    const page = await TeamBookingEmbedPage(input);
    expect(mocks.buildLegacyCtx).toHaveBeenCalledWith({}, {}, params, query);
    expect(mocks.getServerSideProps).toHaveBeenCalledWith({ params, query });
    expect(page.props).toEqual({ eventData: { interfaceLanguage: null }, isEmbed: true });
    expect(await generateMetadata()).toEqual({ robots: { follow: false, index: false } });
  });

  it("preserves the event's configured interface language", async () => {
    mocks.getServerSideProps.mockResolvedValue({ props: { eventData: { interfaceLanguage: "fr" } } });
    mocks.loadTranslations.mockResolvedValue({ common: { book: "Réserver" } });
    const page = await TeamBookingEmbedPage(input);
    expect(mocks.loadTranslations).toHaveBeenCalledWith("fr", "common");
    expect(page.props.locale).toBe("fr");
    expect(page.props.children.props.isEmbed).toBe(true);
  });

  it("keeps missing or inaccessible events unavailable", async () => {
    mocks.getServerSideProps.mockResolvedValue({ notFound: true });
    await expect(TeamBookingEmbedPage(input)).rejects.toThrow("NEXT_NOT_FOUND");
  });

  it("retains embed namespace and layout when the team loader redirects", async () => {
    mocks.getServerSideProps.mockResolvedValue({
      redirect: { destination: "/team/renamed/advisor-discovery", permanent: false },
    });
    await expect(TeamBookingEmbedPage(input)).rejects.toThrow(
      "NEXT_REDIRECT:/team/renamed/advisor-discovery/embed?layout=month_view&embed=wnc-booking"
    );
  });
});
