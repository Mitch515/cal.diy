/**
 * Non-organization replacements for the org/team helpers that used to live under
 * `packages/features/ee`. That directory is covered by Cal.com's commercial
 * license and is not present in this repository, so these are independent
 * implementations written against the call sites in the team booking route.
 *
 * This deployment has no organizations: every team is top-level (`parentId` is
 * null and `isOrganization` is false). Each org-aware helper therefore reduces
 * to its non-org branch, and the team lookups become plain Prisma queries.
 */
import { WEBAPP_URL } from "@calcom/lib/constants";
import prisma from "@calcom/prisma";

/** Prisma `where` fragment selecting a team by slug. */
export const getSlugOrRequestedSlug = (slug: string) => ({ slug });

/**
 * Booking origin for a team page. With no organizations there are no per-org
 * subdomains, so every team is served from the main app origin.
 */
export const getOrgFullOrigin = (_orgSlug: string | null) => WEBAPP_URL;

/**
 * Resolves whether a request arrived on an organization subdomain. Always false
 * here, which keeps the route on its non-org path.
 */
export const orgDomainConfig = (
  _req: unknown,
  _fallback?: string | string[] | null
) => ({
  currentOrgDomain: null as string | null,
  isValidOrgDomain: false,
});

/** SEO settings are an organization-level feature; default to indexable. */
export const getOrganizationSEOSettings = (_team: unknown) => ({
  allowSEOIndexing: true,
});

const teamDataSelect = {
  id: true,
  name: true,
  slug: true,
  logoUrl: true,
  bannerUrl: true,
  isPrivate: true,
  hideBranding: true,
  theme: true,
  brandColor: true,
  darkBrandColor: true,
  parent: {
    select: {
      id: true,
      slug: true,
      name: true,
      logoUrl: true,
      bannerUrl: true,
      hideBranding: true,
      theme: true,
      brandColor: true,
      darkBrandColor: true,
    },
  },
} as const;

/** Team record backing a public team booking page. */
export async function getTeamData(teamSlug: string, _orgSlug: string | null) {
  return prisma.team.findFirst({
    where: { slug: teamSlug, parentId: null },
    select: teamDataSelect,
  });
}

export type TeamData = Awaited<ReturnType<typeof getTeamData>>;

/** Team id for a slug, used to scope event-type and slot queries. */
export async function findTeamIdBySlug(teamSlug: string, _orgSlug: string | null) {
  const team = await prisma.team.findFirst({
    where: { slug: teamSlug, parentId: null },
    select: { id: true },
  });
  return team?.id ?? null;
}
