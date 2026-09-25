// @vitest-environment node
import { createHash } from "node:crypto";
import { expect, test } from "vitest";
import { z } from "zod";
import {
  deliverWncConfirmation,
  type WncConfirmationDependencies,
  type WncConfirmationInput,
  type WncMailState,
  sameWncMailState,
  wncConfirmationContent,
  wncMailStateSchema,
} from "./wnc-confirmation";

const input: WncConfirmationInput = {
  uid: "native-test",
  title: "Client acquisition discussion with Jamie",
  start: "2030-09-14T18:00:00.000Z",
  meetingUrl: "https://teams.microsoft.com/l/meetup-join/demo",
  timeZone: "America/Los_Angeles",
  attendee: { name: "Jamie", email: "jamie@example.org" },
  organizer: { name: "Quentin", email: "quentin@getwealthnavigator.com" },
  participants: [{ name: "Louay", email: "louay@getwealthnavigator.com" }],
  setterEmail: "setter@getwealthnavigator.com",
  eligible: true,
};
const draftSchema = z.object({
  subject: z.string(),
  body: z.object({ contentType: z.string(), content: z.string() }),
  toRecipients: z.array(z.object({ emailAddress: z.object({ address: z.string() }) })),
  ccRecipients: z.array(z.object({ emailAddress: z.object({ address: z.string() }) })),
});
function harness(timeout = false) {
  let state: WncMailState | undefined;
  let message: (z.infer<typeof draftSchema> & { id: string; isDraft: boolean }) | undefined;
  let sends = 0;
  let drafts = 0;
  const deps: WncConfirmationDependencies = {
    async load() {
      return state;
    },
    async save(previous, next) {
      if (JSON.stringify(previous) !== JSON.stringify(state)) return false;
      state = structuredClone(next);
      return true;
    },
    async request(path, method, body) {
      if (path.endsWith("/send")) {
        sends += 1;
        if (timeout) throw new Error("lost response");
        if (message) message.isDraft = false;
        return new Response(null, { status: 202 });
      }
      if (path === "/messages" && method === "POST") {
        drafts += 1;
        message = { ...draftSchema.parse(body), id: "draft-1", isDraft: true };
        return Response.json(message);
      }
      if (path.startsWith("/messages?")) return Response.json({ value: message ? [message] : [] });
      return Response.json(message);
    },
  };
  return {
    deps,
    sends: () => sends,
    drafts: () => drafts,
    markSent() {
      if (message) message.isDraft = false;
    },
    tamper() {
      if (message) message.ccRecipients = [];
    },
  };
}
test("fit and discovery keep To, visible CC and attendance separate", () => {
  expect(wncConfirmationContent({ ...input, participants: [] })).toMatchObject({
    to: "jamie@example.org",
    cc: ["setter@getwealthnavigator.com"],
  });
  expect(
    wncConfirmationContent({
      ...input,
      participants: [...input.participants, input.organizer, input.attendee, ...input.participants],
    })
  ).toMatchObject({
    from: "quentin@getwealthnavigator.com",
    to: "jamie@example.org",
    cc: ["louay@getwealthnavigator.com", "setter@getwealthnavigator.com"],
    subject: input.title,
  });
  expect(input.participants.some((person) => person.email === input.setterEmail)).toBe(false);
  expect(wncConfirmationContent(input).body).toContain("11:00 AM PDT");
});
test("one Microsoft submission survives repeated confirmation recovery", async () => {
  const h = harness();
  expect((await deliverWncConfirmation(input, h.deps)).state).toBe("submitted");
  expect((await deliverWncConfirmation(input, h.deps)).state).toBe("submitted");
  expect(h.sends()).toBe(1);
  expect(h.drafts()).toBe(1);
});
test("an uncertain send is never repeated and can reconcile a later Sent copy", async () => {
  const h = harness(true);
  expect((await deliverWncConfirmation(input, h.deps)).state).toBe("uncertain");
  expect((await deliverWncConfirmation(input, h.deps)).state).toBe("uncertain");
  expect(h.sends()).toBe(1);
  h.markSent();
  expect((await deliverWncConfirmation(input, h.deps)).state).toBe("submitted");
  expect(h.sends()).toBe(1);
});
test("a draft with altered recipients cannot be sent", async () => {
  const h = harness();
  const request = h.deps.request;
  h.deps.request = async (path, method, body) => {
    if (path.startsWith("/messages/draft-1?")) h.tamper();
    return request(path, method, body);
  };
  await expect(deliverWncConfirmation(input, h.deps)).rejects.toThrow("changed");
  expect(h.sends()).toBe(0);
});

test("a claim reloaded with reordered keys still advances to one sent confirmation", async () => {
  // Postgres jsonb and zod both reorder keys, so the stored claim never matches the in-memory one as a string.
  const h = harness();
  let stored: unknown;
  h.deps.load = async () => wncMailStateSchema.optional().parse(stored && JSON.parse(JSON.stringify(stored)));
  h.deps.save = async (previous, next) => {
    const current = await h.deps.load();
    if (!sameWncMailState(current, previous)) return false;
    stored = Object.fromEntries(Object.entries(next).reverse());
    return true;
  };
  expect((await deliverWncConfirmation(input, h.deps)).state).toBe("submitted");
  expect(h.sends()).toBe(1);
  expect(h.drafts()).toBe(1);
});

test("concurrent attempts claim one confirmation and one draft", async () => {
  const h = harness();
  await Promise.all([deliverWncConfirmation(input, h.deps), deliverWncConfirmation(input, h.deps)]);
  expect(h.sends()).toBe(1);
  expect(h.drafts()).toBe(1);
});

test("a booking made before the release never sends, even with a stuck claim and an existing draft", async () => {
  const h = harness();
  h.deps.load = async () => ({ key: "stuck", state: "preparing", claimedAt: "2026-09-24T18:21:05.000Z" });
  await expect(deliverWncConfirmation({ ...input, eligible: false }, h.deps)).rejects.toThrow("predates");
  expect(h.drafts()).toBe(0);
  expect(h.sends()).toBe(0);
});

test("a meeting that has started is never confirmed, even from a stuck claim", async () => {
  const h = harness();
  await expect(
    deliverWncConfirmation({ ...input, start: "2026-09-23T14:00:00.000Z" }, h.deps)
  ).rejects.toThrow("has started");
  expect(h.drafts()).toBe(0);
  expect(h.sends()).toBe(0);
});

const composed = {
  subject: "Our call on Monday, September 14 at 11:00 AM Pacific",
  body: `Hi Jamie,

Personal opening.

You can join us on Microsoft Teams here:
${input.meetingUrl}

Thank you,
Quentin`,
};
test("a WN-composed message is drafted verbatim with Cal's own recipients", async () => {
  const h = harness();
  const drafted: unknown[] = [];
  const request = h.deps.request;
  h.deps.request = async (path, method, body) => {
    if (path === "/messages" && method === "POST") drafted.push(body);
    return request(path, method, body);
  };
  const state = await deliverWncConfirmation({ ...input, message: composed }, h.deps);
  expect(state).toMatchObject({ state: "submitted", subject: composed.subject, body: composed.body });
  expect(draftSchema.parse(drafted[0])).toMatchObject({
    subject: composed.subject,
    body: { contentType: "text", content: composed.body },
    toRecipients: [{ emailAddress: { address: "jamie@example.org" } }],
  });
});
test("a retry keeps the text of the first claim and never sends twice", async () => {
  const h = harness(true);
  await deliverWncConfirmation({ ...input, message: composed }, h.deps);
  h.markSent();
  const retried = await deliverWncConfirmation({ ...input, message: { ...composed, body: `Other ${composed.body}` } }, h.deps);
  expect(retried).toMatchObject({ state: "submitted", body: composed.body });
  expect(h.sends()).toBe(1);
  expect(h.drafts()).toBe(1);
});
test("a claim without saved text keeps the standard wording when a message arrives later", async () => {
  const h = harness();
  const drafted: unknown[] = [];
  const request = h.deps.request;
  h.deps.request = async (path, method, body) => {
    if (path === "/messages" && method === "POST") drafted.push(body);
    return request(path, method, body);
  };
  const key = createHash("sha256").update(`${input.uid}:initial-confirmation`).digest("hex");
  await h.deps.save(undefined, { key, state: "preparing", claimedAt: "2026-01-01T00:00:00.000Z" });
  expect((await deliverWncConfirmation({ ...input, message: composed }, h.deps)).state).toBe("submitted");
  expect(draftSchema.parse(drafted[0]).body.content).toBe(wncConfirmationContent(input).body);
  expect(h.sends()).toBe(1);
});
