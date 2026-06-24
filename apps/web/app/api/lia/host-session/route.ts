import { createHmac, randomUUID } from "crypto";
import { OAuth2Client } from "googleapis-common";
import { encode } from "next-auth/jwt";
import type { NextRequest } from "next/server";
import { NextResponse } from "next/server";

import { getGoogleAppKeys } from "@calcom/app-store/googlecalendar/lib/getGoogleAppKeys";
import {
  GOOGLE_CALENDAR_SCOPES,
  SCOPE_USERINFO_PROFILE,
  WEBAPP_URL,
  WEBAPP_URL_FOR_OAUTH,
} from "@calcom/lib/constants";
import prisma from "@calcom/prisma";

/**
 * LIA host onboarding bridge.
 *
 * A LIA-provisioned booking host has no password and cannot log into Cal.diy,
 * so they cannot reach the "Connect Google/Outlook" screen on their own. LIA
 * mints a single-use token (see /api/lia/provision-host) and sends the host a
 * link that lands here. This route consumes the token, establishes an
 * authenticated NextAuth session for exactly that host, and then sends them
 * STRAIGHT to the calendar provider's OAuth consent screen — skipping the
 * Cal.diy app shell, the app store, and the "install / connect your first
 * calendar" clicks. The host just sees Google's "Allow" screen.
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
  const provider = (req.nextUrl.searchParams.get("provider") || "google").toLowerCase();

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

  // Send them straight to the provider's OAuth consent screen when we can.
  const target = await resolveConnectTarget(provider, user.id, secret);

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
 * Build the destination URL. For Google we generate the OAuth consent URL
 * inline (mirroring googlecalendar/api/add.ts), with the nonce-signed `state`
 * the callback requires, and a same-origin `returnTo` success page so the host
 * never lands in the Cal.diy app shell. Anything else (or any failure — e.g.
 * the app keys not being configured) falls back to the standard connect screen.
 */
async function resolveConnectTarget(
  provider: string,
  userId: number,
  secret: string
): Promise<string> {
  const fallback = new URL("/apps/installed/calendar", WEBAPP_URL).toString();
  if (provider !== "google") return fallback;
  try {
    const { client_id, client_secret } = await getGoogleAppKeys();
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
    const redirect_uri = `${WEBAPP_URL_FOR_OAUTH}/api/integrations/googlecalendar/callback`;
    const oAuth2Client = new OAuth2Client(client_id, client_secret, redirect_uri);
    return oAuth2Client.generateAuthUrl({
      access_type: "offline",
      scope: [SCOPE_USERINFO_PROFILE, ...GOOGLE_CALENDAR_SCOPES],
      prompt: "consent",
      state,
    });
  } catch {
    return fallback;
  }
}
