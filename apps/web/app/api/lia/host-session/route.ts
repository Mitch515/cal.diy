import { randomUUID } from "crypto";
import { encode } from "next-auth/jwt";
import type { NextRequest } from "next/server";
import { NextResponse } from "next/server";

import { WEBAPP_URL } from "@calcom/lib/constants";
import prisma from "@calcom/prisma";

/**
 * LIA host onboarding bridge.
 *
 * A LIA-provisioned booking host has no password and cannot log into Cal.diy,
 * so they cannot reach the "Connect Google/Outlook" screen on their own. LIA
 * mints a single-use token (see /api/lia/provision-host) and sends the host a
 * link that lands here. This route consumes the token, establishes an
 * authenticated NextAuth session for exactly that host, and redirects them to
 * the calendar-connect screen.
 *
 * IMPORTANT: this endpoint is BROWSER-FACING and must NOT require the internal
 * provisioning secret header — a browser navigation cannot set headers. The
 * single-use, short-lived VerificationToken IS the credential. The internal
 * secret only guards the server-to-server minting endpoint.
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
      // Best-effort cleanup of an expired / malformed record.
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
      // Single-profile users use the `usr-<id>` upId convention in this fork.
      upId: `usr-${user.id}`,
      locale: user.locale ?? "en",
      belongsToActiveTeam: false,
      // Distinguishes this minted session in logs; harmless if unused.
      jti: randomUUID(),
    },
  });

  const useSecureCookies = WEBAPP_URL.startsWith("https://");
  const cookiePrefix = useSecureCookies ? "__Secure-" : "";
  const cookieName = `${cookiePrefix}next-auth.session-token`;

  const res = NextResponse.redirect(new URL("/apps/installed/calendar", WEBAPP_URL), { status: 302 });
  res.cookies.set(cookieName, sessionToken, {
    httpOnly: true,
    secure: useSecureCookies,
    sameSite: useSecureCookies ? "none" : "lax",
    path: "/",
    maxAge: SESSION_MAX_AGE_SECONDS,
    domain: process.env.NEXTAUTH_COOKIE_DOMAIN || undefined,
  });
  return res;
}
