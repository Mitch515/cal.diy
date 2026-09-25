import { createHash } from "node:crypto";
import process from "node:process";
import { ErrorCode } from "@calcom/lib/errorCodes";
import { ErrorWithCode } from "@calcom/lib/errors";
import prisma from "@calcom/prisma";
import { z } from "zod";

const ADDRESS = "quentin@getwealthnavigator.com";
const PROPERTY = "String {5c5b1cf8-cd3b-42fb-9103-d59ae1ec2934} Name WncConfirmation";
export const wncMailStateSchema = z.object({
  state: z.enum(["preparing", "draft", "sending", "submitted", "uncertain"]),
  key: z.string(),
  draftId: z.string().optional(),
  claimedAt: z.string().datetime().optional(),
});
const metadataSchema = z.object({
  wncSetterEmail: z.string().email(),
  wncAttendeeEmail: z.string().email(),
  wncAttendeeTimeZone: z.string(),
  wncConfirmation: wncMailStateSchema.optional(),
});
const messageSchema = z.object({
  id: z.string(),
  isDraft: z.boolean(),
  toRecipients: z.array(z.object({ emailAddress: z.object({ address: z.string() }) })),
  ccRecipients: z.array(z.object({ emailAddress: z.object({ address: z.string() }) })),
  subject: z.string(),
  body: z.object({ content: z.string(), contentType: z.string() }),
});
export type WncMailState = z.infer<typeof wncMailStateSchema>;
/** Field comparison: Postgres jsonb and zod reorder keys, so JSON.stringify never matches a reloaded claim. */
export function sameWncMailState(left: WncMailState | undefined, right: WncMailState | undefined) {
  if (!left || !right) return left === right;
  return (
    left.state === right.state &&
    left.key === right.key &&
    left.draftId === right.draftId &&
    left.claimedAt === right.claimedAt
  );
}
export interface WncConfirmationInput {
  uid: string;
  title: string;
  start: string;
  meetingUrl: string;
  timeZone: string;
  attendee: { name: string; email: string };
  organizer: { name: string; email: string };
  participants: { name: string; email: string }[];
  setterEmail: string;
}
export function wncConfirmationContent(input: WncConfirmationInput) {
  const normalized = (value: string) => value.trim().toLowerCase();
  if (
    normalized(input.organizer.email) !== ADDRESS ||
    !input.setterEmail.endsWith("@getwealthnavigator.com")
  ) {
    throw new ErrorWithCode(
      ErrorCode.InternalServerError,
      "WNC confirmation requires the approved organizer and verified setter"
    );
  }
  const to = normalized(input.attendee.email);
  const cc = Array.from(
    new Set([...input.participants.map((person) => normalized(person.email)), normalized(input.setterEmail)])
  )
    .filter((email) => email !== ADDRESS && email !== to)
    .sort();
  const time = (zone: string) =>
    new Intl.DateTimeFormat("en-US", {
      weekday: "long",
      month: "long",
      day: "numeric",
      hour: "numeric",
      minute: "2-digit",
      timeZone: zone,
      timeZoneName: "short",
    }).format(new Date(input.start));
  const body = [
    `Hi ${input.attendee.name},`,
    "",
    `Confirming ${input.title} on ${time("America/New_York")}.`,
    ...(input.timeZone !== "America/New_York" ? [`That is ${time(input.timeZone)} in your time zone.`] : []),
    "",
    `Join on Microsoft Teams: ${input.meetingUrl}`,
    "",
    "Please reply if you need to change the time.",
    "",
    "Thank you,",
    "Quentin",
  ].join("\n");
  return { from: ADDRESS, to, cc, subject: input.title, body };
}
function setting(key: string) {
  const value = Reflect.get(process.env, key);
  return typeof value === "string" ? value.trim() : "";
}
export function wncConfirmationConfigured() {
  return ["WNC_MICROSOFT_TENANT_ID", "WNC_MICROSOFT_CLIENT_ID", "WNC_MICROSOFT_CLIENT_SECRET"].every((key) =>
    Boolean(setting(key))
  );
}
async function graphRequest(path: string, method = "GET", body?: object) {
  const tokenResponse = await fetch(
    `https://login.microsoftonline.com/${encodeURIComponent(setting("WNC_MICROSOFT_TENANT_ID"))}/oauth2/v2.0/token`,
    {
      method: "POST",
      signal: AbortSignal.timeout(15_000),
      body: new URLSearchParams({
        client_id: setting("WNC_MICROSOFT_CLIENT_ID"),
        client_secret: setting("WNC_MICROSOFT_CLIENT_SECRET"),
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
      }),
    }
  );
  if (!tokenResponse.ok)
    throw new ErrorWithCode(ErrorCode.InternalServerError, "Microsoft confirmation authentication failed");
  const token = z.object({ access_token: z.string() }).parse(await tokenResponse.json());
  return fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(ADDRESS)}${path}`, {
    method,
    redirect: "error",
    signal: AbortSignal.timeout(20_000),
    headers: {
      Authorization: `Bearer ${token.access_token}`,
      "Content-Type": "application/json",
      Prefer: 'IdType="ImmutableId"',
    },
    ...(body ? { body: JSON.stringify(body) } : {}),
  });
}
export interface WncConfirmationDependencies {
  load(): Promise<WncMailState | undefined>;
  save(previous: WncMailState | undefined, next: WncMailState): Promise<boolean>;
  request(path: string, method?: string, body?: object): Promise<Response>;
}

/** Persist before sending. An uncertain send is only reconciled, never blindly repeated. */
export async function deliverWncConfirmation(
  input: WncConfirmationInput,
  deps: WncConfirmationDependencies
): Promise<WncMailState> {
  const content = wncConfirmationContent(input);
  const key = createHash("sha256").update(`${input.uid}:initial-confirmation`).digest("hex");
  let state = await deps.load();
  if (state?.state === "submitted") return state;
  // A claim stuck before this fix can be retried later; never confirm a meeting that has already started.
  if (Date.parse(input.start) <= Date.now())
    throw new ErrorWithCode(ErrorCode.BadRequest, "The meeting has started; its confirmation is not sent");
  if (state?.state === "preparing" && Date.now() - Date.parse(state.claimedAt ?? "1970-01-01") < 180_000)
    return state;
  if (!state) {
    const initial: WncMailState = { key, state: "preparing", claimedAt: new Date().toISOString() };
    if (!(await deps.save(undefined, initial))) return (await deps.load()) ?? initial;
    state = initial;
  } else if (state.state === "preparing") {
    const claimed: WncMailState = { ...state, claimedAt: new Date().toISOString() };
    if (!(await deps.save(state, claimed))) return (await deps.load()) ?? claimed;
    state = claimed;
  }
  if (state.key !== key)
    throw new ErrorWithCode(ErrorCode.InternalServerError, "Confirmation identity mismatch");
  const query = new URLSearchParams({
    $filter: `singleValueExtendedProperties/Any(p: p/id eq '${PROPERTY}' and p/value eq '${key}')`,
    $select: "id,isDraft,toRecipients,ccRecipients,subject,body",
    $top: "2",
  });
  const existing = await deps.request(`/messages?${query}`);
  if (!existing.ok)
    throw new ErrorWithCode(ErrorCode.InternalServerError, "Cannot reconcile confirmation mail");
  const messages = z.object({ value: z.array(messageSchema) }).parse(await existing.json()).value;
  if (messages.length > 1)
    throw new ErrorWithCode(ErrorCode.InternalServerError, "Duplicate confirmation drafts need review");
  let message = messages[0];
  if (state.state === "sending" || state.state === "uncertain") {
    const recovered: WncMailState = {
      ...state,
      state: message && !message.isDraft ? "submitted" : "uncertain",
    };
    await deps.save(state, recovered);
    return recovered;
  }
  if (!message) {
    const created = await deps.request("/messages", "POST", {
      subject: content.subject,
      body: { contentType: "text", content: content.body },
      toRecipients: [{ emailAddress: { address: content.to } }],
      ccRecipients: content.cc.map((address) => ({ emailAddress: { address } })),
      replyTo: [{ emailAddress: { address: ADDRESS } }],
      singleValueExtendedProperties: [{ id: PROPERTY, value: key }],
    });
    if (!created.ok)
      throw new ErrorWithCode(ErrorCode.InternalServerError, "Confirmation draft creation failed");
    message = messageSchema.parse(await created.json());
  }
  const readBack = await deps.request(
    `/messages/${encodeURIComponent(message.id)}?$select=id,isDraft,toRecipients,ccRecipients,subject,body`
  );
  if (!readBack.ok)
    throw new ErrorWithCode(ErrorCode.InternalServerError, "Confirmation draft could not be read back");
  message = messageSchema.parse(await readBack.json());
  const recipientList = (recipients: typeof message.toRecipients) =>
    recipients.map((value) => value.emailAddress.address.toLowerCase()).sort();
  if (
    JSON.stringify(recipientList(message.toRecipients)) !== JSON.stringify([content.to]) ||
    JSON.stringify(recipientList(message.ccRecipients)) !== JSON.stringify(content.cc) ||
    message.subject !== content.subject ||
    message.body.content !== content.body ||
    message.body.contentType.toLowerCase() !== "text"
  ) {
    throw new ErrorWithCode(ErrorCode.InternalServerError, "Confirmation draft changed and needs review");
  }
  if (!message.isDraft) {
    const submitted: WncMailState = { key, state: "submitted", draftId: message.id };
    await deps.save(state, submitted);
    return submitted;
  }
  const sending: WncMailState = { key, state: "sending", draftId: message.id };
  if (!(await deps.save(state, sending))) return (await deps.load()) ?? sending;
  try {
    const response = await deps.request(`/messages/${encodeURIComponent(message.id)}/send`, "POST");
    const result: WncMailState = { ...sending, state: response.status === 202 ? "submitted" : "uncertain" };
    await deps.save(sending, result);
    return result;
  } catch {
    const uncertain: WncMailState = { ...sending, state: "uncertain" };
    await deps.save(sending, uncertain);
    return uncertain;
  }
}

export async function submitWncConfirmation(input: WncConfirmationInput) {
  return deliverWncConfirmation(input, {
    request: graphRequest,
    async load() {
      const booking = await prisma.booking.findUniqueOrThrow({
        where: { uid: input.uid },
        select: { metadata: true },
      });
      return metadataSchema.parse(booking.metadata).wncConfirmation;
    },
    async save(previous, next) {
      const booking = await prisma.booking.findUniqueOrThrow({
        where: { uid: input.uid },
        select: { id: true, metadata: true, updatedAt: true },
      });
      const metadata = metadataSchema.parse(booking.metadata);
      if (!sameWncMailState(metadata.wncConfirmation, previous)) return false;
      if (!booking.metadata || typeof booking.metadata !== "object" || Array.isArray(booking.metadata))
        return false;
      const saved = await prisma.booking.updateMany({
        where: { id: booking.id, updatedAt: booking.updatedAt },
        data: { metadata: { ...booking.metadata, wncConfirmation: next } },
      });
      return saved.count === 1;
    },
  });
}

export async function isWncNativeBooking(uid?: string | null) {
  let current = uid;
  for (let depth = 0; current && depth < 20; depth += 1) {
    const booking = await prisma.booking.findUnique({
      where: { uid: current },
      select: { idempotencyKey: true, fromReschedule: true, metadata: true },
    });
    if (booking?.idempotencyKey?.startsWith("wnc:")) return true;
    if (z.object({ wncRequestId: z.string().regex(/^[a-f0-9]{64}$/) }).safeParse(booking?.metadata).success)
      return true;
    current = booking?.fromReschedule ?? undefined;
  }
  return false;
}
