import { Router } from "express";
import {
  getUser,
  upsertUser,
  createResetToken,
  consumeResetToken,
  logAppError,
} from "../auth/app-store";
import {
  hashPassword,
  verifyPassword,
  signSessionToken,
  requireAuth,
  SESSION_COOKIE_NAME,
} from "../auth/app-auth";
import { sendSystemEmail } from "../tools/system-email-service";

const RESET_TOKEN_TTL_MS = 60 * 60 * 1000; // 1 hour
const PUBLIC_BASE_URL = process.env.PUBLIC_BASE_URL || "http://localhost:3000";

function log(msg: string) {
  process.stderr.write(`[agent365-bridge] ${msg}\n`);
}

function cookieOptions() {
  return {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax" as const,
    maxAge: 8 * 60 * 60 * 1000,
    path: "/",
  };
}

export function createAuthRoutes(): Router {
  const router = Router();

  router.post("/login", async (req, res) => {
    const { email, password } = req.body ?? {};
    if (typeof email !== "string" || typeof password !== "string") {
      res.status(400).json({ error: "email and password are required" });
      return;
    }

    const user = await getUser(email);
    if (!user || !user.isActive) {
      res.status(401).json({ error: "invalid_credentials" });
      return;
    }

    const valid = await verifyPassword(password, user.passwordHash);
    if (!valid) {
      res.status(401).json({ error: "invalid_credentials" });
      return;
    }

    const token = signSessionToken({ email: user.email, role: user.role });
    res.cookie(SESSION_COOKIE_NAME, token, cookieOptions());
    res.json({ email: user.email, role: user.role });
  });

  router.post("/logout", (_req, res) => {
    res.clearCookie(SESSION_COOKIE_NAME, { path: "/" });
    res.json({ ok: true });
  });

  router.get("/me", requireAuth, (req, res) => {
    res.json(req.appUser);
  });

  router.post("/forgot-password", async (req, res) => {
    const { email } = req.body ?? {};
    if (typeof email !== "string") {
      res.status(400).json({ error: "email is required" });
      return;
    }

    // Always respond 200 to avoid leaking whether an account exists.
    res.json({ ok: true });

    try {
      const user = await getUser(email);
      if (!user || !user.isActive) return;

      const token = await createResetToken(user.email, RESET_TOKEN_TTL_MS);
      const link = `${PUBLIC_BASE_URL}/app/reset-password?token=${encodeURIComponent(token)}`;
      await sendSystemEmail(
        user.email,
        "Reset your password",
        `<p>A password reset was requested for your account.</p>` +
          `<p><a href="${link}">Click here to set a new password</a>. This link expires in 1 hour.</p>` +
          `<p>If you didn't request this, you can ignore this email.</p>`
      );
      log(`Password reset email sent to ${user.email}`);
    } catch (err) {
      log(`forgot-password email send failed: ${err}`);
      await logAppError("forgot-password", "Failed to send password reset email", String(err));
    }
  });

  router.post("/reset-password", async (req, res) => {
    const { token, newPassword } = req.body ?? {};
    if (typeof token !== "string" || typeof newPassword !== "string" || newPassword.length < 8) {
      res.status(400).json({ error: "token and a password of at least 8 characters are required" });
      return;
    }

    const email = await consumeResetToken(token);
    if (!email) {
      res.status(400).json({ error: "invalid_or_expired_token" });
      return;
    }

    const user = await getUser(email);
    if (!user) {
      res.status(400).json({ error: "invalid_or_expired_token" });
      return;
    }

    await upsertUser({ ...user, passwordHash: await hashPassword(newPassword) });
    res.json({ ok: true });
  });

  return router;
}
