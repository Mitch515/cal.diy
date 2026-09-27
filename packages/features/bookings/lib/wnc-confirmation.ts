import { createHash } from "node:crypto";
import process from "node:process";
import { ErrorCode } from "@calcom/lib/errorCodes";
import { ErrorWithCode } from "@calcom/lib/errors";
import prisma from "@calcom/prisma";
import { z } from "zod";

const ADDRESS = "quentin@getwealthnavigator.com";
const PROPERTY = "String {5c5b1cf8-cd3b-42fb-9103-d59ae1ec2934} Name WncConfirmation";
/** An image referenced from an HTML body as `cid:<contentId>`, such as Quentin's signature. */
export const wncInlineImageSchema = z.object({
  contentId: z.string().regex(/^[A-Za-z0-9._-]{1,64}$/),
  name: z.string().min(1).max(100),
  contentType: z.enum(["image/jpeg", "image/png"]),
  contentBase64: z.string().min(1).max(400_000),
});
export type WncInlineImage = z.infer<typeof wncInlineImageSchema>;
export const wncMailStateSchema = z.object({
  state: z.enum(["preparing", "draft", "sending", "submitted", "uncertain"]),
  key: z.string(),
  draftId: z.string().optional(),
  claimedAt: z.string().datetime().optional(),
  /** Saved with the first claim so every retry drafts, compares and sends the same text. */
  subject: z.string().optional(),
  body: z.string().optional(),
  contentType: z.enum(["text", "html"]).optional(),
  inlineImages: z.array(wncInlineImageSchema).max(2).optional(),
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
  attachments: z.array(z.object({ contentId: z.string().nullish(), isInline: z.boolean().optional() })).optional(),
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
  /** False for every booking made before the September 25 release. Those claims were stuck by the key-order bug and are never sent. */
  eligible: boolean;
  /** WN's composed confirmation. Recipients stay Cal's; WN supplies the text. */
  message?: { subject: string; body: string; contentType?: "html"; inlineImages?: WncInlineImage[] };
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
  if (input.message) {
    const { subject, body: composed, inlineImages = [] } = input.message;
    const contentType: "text" | "html" = input.message.contentType ?? "text";
    return { from: ADDRESS, to, cc, subject, body: composed, contentType, inlineImages };
  }
  return { from: ADDRESS, to, cc, subject: input.title, body, contentType: "text" as const, inlineImages: [] };
}
type ClaimedMessage = Pick<WncMailState, "subject" | "body" | "contentType" | "inlineImages">;
/** The message of an existing claim wins over whatever a later retry supplies. */
function claimedMessage(state: WncMailState | undefined): ClaimedMessage {
  if (state?.subject === undefined || state.body === undefined) return {};
  return { subject: state.subject, body: state.body, contentType: state.contentType ?? "text", inlineImages: state.inlineImages ?? [] };
}
const ENTITIES: Record<string, string> = { "&nbsp;": " ", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&#39;": "'", "&amp;": "&" };
const decodeHtml = (text: string) => text.replace(/&(?:nbsp|lt|gt|quot|#39|amp);/g, (entity) => ENTITIES[entity] ?? entity);
/** Values of one attribute across one tag, for example every link's href. Graph keeps double quotes but reorders attributes. */
function attributeValues(html: string, tag: string, attribute: string) {
  const tags = html.match(new RegExp(`<${tag}\\b[^>]*>`, "gi")) ?? [];
  return tags
    .map((element) => new RegExp(`\\b${attribute}="([^"]*)"`, "i").exec(element)?.[1])
    .filter((value): value is string => value !== undefined)
    .map(decodeHtml)
    .sort();
}
function visibleText(html: string) {
  const withoutHead = html.replace(/<head[\s\S]*?<\/head>/gi, "").replace(/<(style|script)[\s\S]*?<\/\1>/gi, "");
  return decodeHtml(withoutHead.replace(/<[^>]*>/g, " ")).replace(/\s+/g, " ").trim();
}
/**
 * Graph rewrites an HTML draft when it is read back (it adds a head and styles and reorders attributes),
 * so HTML drafts are compared by what the reader sees and where the links and images point.
 */
export function sameHtmlContent(left: string, right: string) {
  const same = (tag: string, attribute: string) =>
    JSON.stringify(attributeValues(left, tag, attribute)) === JSON.stringify(attributeValues(right, tag, attribute));
  return visibleText(left) === visibleText(right) && same("a", "href") && same("img", "src");
}
/** WN's composed body must point at the booking's own Teams meeting. */
export function bodyLinksTo(body: string, contentType: "text" | "html" | undefined, url: string) {
  return contentType === "html" ? attributeValues(body, "a", "href").includes(url) : body.includes(url);
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

type Claim = { proceed: false; state: WncMailState } | { proceed: true; state: WncMailState; claimedNow: boolean };
type Content = ReturnType<typeof wncConfirmationContent>;
type Message = z.infer<typeof messageSchema>;

/** Claim the confirmation, or return the state another attempt already holds. */
async function claimConfirmation(
  deps: WncConfirmationDependencies,
  key: string,
  start: string,
  text: ClaimedMessage
): Promise<Claim> {
  const state = await deps.load();
  if (state?.state === "submitted") return { proceed: false, state };
  // A claim stuck before the key-order fix can be retried later; never confirm a meeting that has already started.
  if (Date.parse(start) <= Date.now())
    throw new ErrorWithCode(ErrorCode.BadRequest, "The meeting has started; its confirmation is not sent");
  if (state?.state === "preparing" && Date.now() - Date.parse(state.claimedAt ?? "1970-01-01") < 180_000)
    return { proceed: false, state };
  if (!state) {
    const initial: WncMailState = { key, state: "preparing", claimedAt: new Date().toISOString(), ...text };
    if (!(await deps.save(undefined, initial))) return { proceed: false, state: (await deps.load()) ?? initial };
    return { proceed: true, state: initial, claimedNow: true };
  }
  if (state.state !== "preparing") return { proceed: true, state, claimedNow: false };
  const claimed: WncMailState = { ...state, claimedAt: new Date().toISOString() };
  if (!(await deps.save(state, claimed))) return { proceed: false, state: (await deps.load()) ?? claimed };
  return { proceed: true, state: claimed, claimedNow: false };
}

function bodyMatches(message: Message, content: Content) {
  if (message.body.contentType.toLowerCase() !== content.contentType) return false;
  if (content.contentType === "text") return message.body.content === content.body;
  const attached = new Set(
    (message.attachments ?? []).filter((file) => file.isInline).map((file) => file.contentId?.replace(/^<|>$/g, ""))
  );
  return sameHtmlContent(message.body.content, content.body) && content.inlineImages.every((image) => attached.has(image.contentId));
}

function draftMatches(message: Message, content: Content): boolean {
  const recipientList = (recipients: Message["toRecipients"]) =>
    recipients.map((value) => value.emailAddress.address.toLowerCase()).sort();
  return (
    JSON.stringify(recipientList(message.toRecipients)) === JSON.stringify([content.to]) &&
    JSON.stringify(recipientList(message.ccRecipients)) === JSON.stringify(content.cc) &&
    message.subject === content.subject &&
    bodyMatches(message, content)
  );
}

async function findTaggedMessage(deps: WncConfirmationDependencies, key: string): Promise<Message | undefined> {
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
  return messages[0];
}

async function createDraft(deps: WncConfirmationDependencies, key: string, content: Content): Promise<Message> {
  const created = await deps.request("/messages", "POST", {
    subject: content.subject,
    body: { contentType: content.contentType, content: content.body },
    toRecipients: [{ emailAddress: { address: content.to } }],
    ccRecipients: content.cc.map((address) => ({ emailAddress: { address } })),
    replyTo: [{ emailAddress: { address: ADDRESS } }],
    singleValueExtendedProperties: [{ id: PROPERTY, value: key }],
    ...(content.inlineImages.length > 0
      ? {
          attachments: content.inlineImages.map((image) => ({
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: image.name,
            contentType: image.contentType,
            contentBytes: image.contentBase64,
            contentId: image.contentId,
            isInline: true,
          })),
        }
      : {}),
  });
  if (!created.ok) throw new ErrorWithCode(ErrorCode.InternalServerError, "Confirmation draft creation failed");
  return messageSchema.parse(await created.json());
}

async function readBackDraft(deps: WncConfirmationDependencies, message: Message): Promise<Message> {
  const readBack = await deps.request(
    `/messages/${encodeURIComponent(message.id)}?$select=id,isDraft,toRecipients,ccRecipients,subject,body&$expand=attachments($select=contentId,isInline)`
  );
  if (!readBack.ok)
    throw new ErrorWithCode(ErrorCode.InternalServerError, "Confirmation draft could not be read back");
  return messageSchema.parse(await readBack.json());
}

/** Claim the send before calling Microsoft; a lost response is recorded as uncertain, never retried. */
async function submitDraft(
  deps: WncConfirmationDependencies,
  state: WncMailState,
  identity: ClaimedMessage & { key: string; draftId: string },
  isDraft: boolean
): Promise<WncMailState> {
  if (!isDraft) {
    const submitted: WncMailState = { ...identity, state: "submitted" };
    await deps.save(state, submitted);
    return submitted;
  }
  const sending: WncMailState = { ...identity, state: "sending" };
  if (!(await deps.save(state, sending))) return (await deps.load()) ?? sending;
  try {
    const response = await deps.request(`/messages/${encodeURIComponent(identity.draftId)}/send`, "POST");
    const result: WncMailState = { ...sending, state: response.status === 202 ? "submitted" : "uncertain" };
    await deps.save(sending, result);
    return result;
  } catch {
    const uncertain: WncMailState = { ...sending, state: "uncertain" };
    await deps.save(sending, uncertain);
    return uncertain;
  }
}

/** Persist before sending. An uncertain send is only reconciled, never blindly repeated. */
export async function deliverWncConfirmation(
  input: WncConfirmationInput,
  deps: WncConfirmationDependencies
): Promise<WncMailState> {
  if (!input.eligible)
    throw new ErrorWithCode(ErrorCode.BadRequest, "This booking predates automatic confirmations; confirm it by hand");
  const fresh = wncConfirmationContent(input);
  const key = createHash("sha256").update(`${input.uid}:initial-confirmation`).digest("hex");
  const offered: ClaimedMessage = input.message
    ? { subject: fresh.subject, body: fresh.body, contentType: fresh.contentType, inlineImages: fresh.inlineImages }
    : {};
  const claim = await claimConfirmation(deps, key, input.start, offered);
  if (!claim.proceed) return claim.state;
  const state = claim.state;
  if (state.key !== key)
    throw new ErrorWithCode(ErrorCode.InternalServerError, "Confirmation identity mismatch");
  const text = claimedMessage(state);
  // A claim made before WN supplied text used the standard wording, so a later message cannot replace it.
  const content = { ...(claim.claimedNow ? fresh : wncConfirmationContent({ ...input, message: undefined })), ...text };
  const tagged = await findTaggedMessage(deps, key);
  if (state.state === "sending" || state.state === "uncertain") {
    const recovered: WncMailState = {
      ...state,
      state: tagged && !tagged.isDraft ? "submitted" : "uncertain",
    };
    await deps.save(state, recovered);
    return recovered;
  }
  const message = await readBackDraft(deps, tagged ?? (await createDraft(deps, key, content)));
  if (!draftMatches(message, content)) {
    throw new ErrorWithCode(ErrorCode.InternalServerError, "Confirmation draft changed and needs review");
  }
  return submitDraft(deps, state, { ...text, key, draftId: message.id }, message.isDraft);
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
