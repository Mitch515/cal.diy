import setDefaultConferencingApp from "@calcom/app-store/_utils/setDefaultConferencingApp";
import prisma from "@calcom/prisma";
import type { Prisma } from "@calcom/prisma/client";

export async function ensureMicrosoftTeamsConnection({
  userId,
  key,
}: {
  userId: number;
  key: Prisma.InputJsonObject;
}) {
  const existing = await prisma.credential.findMany({
    where: { userId, appId: "msteams", type: "office365_video" },
    select: { id: true },
    orderBy: { id: "asc" },
  });
  const primary = existing[0];
  const credential = primary
    ? await prisma.credential.update({
        where: { id: primary.id },
        data: { key, invalid: false },
        select: { id: true },
      })
    : await prisma.credential.create({
        data: {
          userId,
          appId: "msteams",
          type: "office365_video",
          key,
        },
        select: { id: true },
      });

  const duplicateIds = existing.slice(1).map(({ id }) => id);
  if (duplicateIds.length > 0) {
    await prisma.credential.deleteMany({
      where: { id: { in: duplicateIds }, userId, type: "office365_video" },
    });
  }

  await setDefaultConferencingApp(userId, "msteams");
  return credential;
}
