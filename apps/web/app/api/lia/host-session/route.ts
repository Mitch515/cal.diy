import { createHmac, randomUUID } from "node:crypto";
import process from "node:process";
import getAppKeysFromSlug from "@calcom/app-store/_utils/getAppKeysFromSlug";
import { getOutlookTenantId } from "@calcom/features/auth/lib/outlook";
import { MICROSOFT_CALENDAR_AND_TEAMS_SCOPES, WEBAPP_URL, WEBAPP_URL_FOR_OAUTH } from "@calcom/lib/constants";
import prisma from "@calcom/prisma";
import type { NextRequest } from "next/server";
import { NextResponse } from "next/server";
import { encode } from "next-auth/jwt";

/**
 * LIA host onboarding bridge.
 *
 * A LIA-provisioned booking host has no password and cannot log into Cal.diy,
 * so they cannot reach the "Connect Outlook" screen on their own. LIA
 * mints a single-use token (see /api/lia/provision-host) and sends the host a
 * link that lands here. This route consumes the token, establishes an
 * authenticated NextAuth session for exactly that host, and then sends them
 * straight to Microsoft's OAuth consent screen, skipping the
 * Cal.diy app shell, the app store, and the "install / connect your first
 * calendar" clicks. One consent connects Outlook Calendar and Teams.
 *
 * IMPORTANT: this endpoint is BROWSER-FACING and must NOT require the internal
 * provisioning secret header — a browser navigation cannot set headers. The
 * single-use, short-lived VerificationToken IS the credential.
 */

const SESSION_MAX_AGE_SECONDS = 60 * 60 * 24 * 30; // 30 days, matches NextAuth default
const TOKEN_PREFIX = "lia-host-session:";

export async function GET(req: NextRequest) {
  const secret = process.env.NEXTAUTH_SECRET;
  if (!secret) {
    return NextResponse.json({ message: "Authentication is not configured" }, { status: 503 });
  }

  const token = req.nextUrl.searchParams.get("token");
  if (!token) {
    return NextResponse.json({ message: "Missing token" }, { status: 400 });
  }
  const record = await prisma.verificationToken.findUnique({ where: { token } });
  const isValid =
    record && record.identifier.startsWith(TOKEN_PREFIX) && record.expires.getTime() > Date.now();
  if (!isValid) {
    if (record) {
      await prisma.verificationToken.delete({ where: { id: record.id } }).catch(() => undefined);
    }
    return NextResponse.json(
      { message: "This calendar setup link is invalid or has expired." },
      { status: 410 }
    );
  }

  const userId = Number(record.identifier.slice(TOKEN_PREFIX.length));
  if (!Number.isInteger(userId) || userId <= 0) {
    await prisma.verificationToken.delete({ where: { id: record.id } }).catch(() => undefined);
    return NextResponse.json({ message: "Invalid calendar setup link." }, { status: 400 });
  }

  const user = await prisma.user.findUnique({
    where: { id: userId },
    select: { id: true, email: true, name: true, username: true, role: true, locale: true },
  });
  if (!user) {
    await prisma.verificationToken.delete({ where: { id: record.id } }).catch(() => undefined);
    return NextResponse.json({ message: "Booking host not found." }, { status: 404 });
  }

  // Single use: consume the token before establishing the session.
  await prisma.verificationToken.delete({ where: { id: record.id } }).catch(() => undefined);

  const sessionToken = await encode({
    secret,
    maxAge: SESSION_MAX_AGE_SECONDS,
    token: {
      sub: String(user.id),
      id: user.id,
      email: user.email,
      name: user.name,
      username: user.username,
      role: user.role,
      upId: `usr-${user.id}`,
      locale: user.locale ?? "en",
      belongsToActiveTeam: false,
      jti: randomUUID(),
    },
  });

  // Send them straight to Microsoft's OAuth consent screen when we can.
  const target = await resolveConnectTarget(user.id, user.email, secret);

  const useSecureCookies = WEBAPP_URL.startsWith("https://");
  const cookiePrefix = useSecureCookies ? "__Secure-" : "";
  const res = NextResponse.redirect(target, { status: 302 });
  res.cookies.set(`${cookiePrefix}next-auth.session-token`, sessionToken, {
    httpOnly: true,
    secure: useSecureCookies,
    sameSite: useSecureCookies ? "none" : "lax",
    path: "/",
    maxAge: SESSION_MAX_AGE_SECONDS,
    domain: process.env.NEXTAUTH_COOKIE_DOMAIN || undefined,
  });
  return res;
}

/**
 * Build Microsoft's OAuth consent URL inline, with the nonce-signed `state`
 * the callback requires, and a same-origin `returnTo` success page so the host
 * never lands in the Cal.diy app shell. A configuration failure falls back to
 * the standard connect screen.
 */
export async function resolveConnectTarget(
  userId: number,
  userEmail: string,
  secret: string,
  getAppKeys: typeof getAppKeysFromSlug = getAppKeysFromSlug
): Promise<string> {
  const fallback = new URL("/apps/installed/calendar", WEBAPP_URL).toString();
  try {
    const appKeys = await getAppKeys("office365-calendar");
    const clientId = typeof appKeys.client_id === "string" ? appKeys.client_id : "";
    if (!clientId) return fallback;
    const nonce = randomUUID();
    const nonceHash = createHmac("sha256", secret).update(`${nonce}:${userId}`).digest("hex");
    const returnTo = new URL("/lia/connected", WEBAPP_URL).toString();
    const state = JSON.stringify({
      returnTo,
      onErrorReturnTo: returnTo,
      fromApp: false,
      nonce,
      nonceHash,
    });
    const params = new URLSearchParams({
      response_type: "code",
      scope: MICROSOFT_CALENDAR_AND_TEAMS_SCOPES.join(" "),
      client_id: clientId,
      prompt: "select_account",
      login_hint: userEmail,
      redirect_uri: `${WEBAPP_URL_FOR_OAUTH}/api/integrations/office365calendar/callback`,
      state,
    });
    return `https://login.microsoftonline.com/${getOutlookTenantId()}/oauth2/v2.0/authorize?${params.toString()}`;
  } catch {
    return fallback;
  }
}
