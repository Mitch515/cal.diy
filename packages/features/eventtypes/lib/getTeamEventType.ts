import { prisma } from "@calcom/prisma";
import { getPublicEventSelect } from "./getPublicEvent";

/**
 * Selects a team by slug. The previous implementation lived under
 * `packages/features/ee` and also resolved unpublished organizations via
 * `metadata.requestedSlug`; this deployment has no organizations, so a plain
 * slug match is equivalent. Defined locally to avoid a package -> app import.
 */
const getSlugOrRequestedSlug = (slug: string) => ({ slug });

export async function getTeamEventType(teamSlug: string, meetingSlug: string, orgSlug: string | null) {
  return await prisma.eventType.findFirst({
    where: {
      team: {
        ...getSlugOrRequestedSlug(teamSlug),
        parent: orgSlug ? getSlugOrRequestedSlug(orgSlug) : null,
      },
      OR: [{ slug: meetingSlug }, { slug: { startsWith: `${meetingSlug}-team-id-` } }],
    },
    // IMPORTANT:
    // This ensures that the queried event type has everything expected in Booker
    select: getPublicEventSelect(false),
    orderBy: {
      slug: "asc",
    },
  });
}
