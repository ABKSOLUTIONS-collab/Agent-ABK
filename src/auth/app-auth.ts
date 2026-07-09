import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";
import type { Request, Response, NextFunction } from "express";
import type { AppRole } from "./app-store";

export const SESSION_COOKIE_NAME = "app_session";
const SESSION_TTL = "8h";

function jwtSecret(): string {
  const secret = process.env.JWT_SECRET;
  if (!secret) {
    throw new Error("JWT_SECRET must be set to use the login/admin app");
  }
  return secret;
}

export interface SessionPayload {
  email: string;
  role: AppRole;
}

declare global {
  // eslint-disable-next-line @typescript-eslint/no-namespace
  namespace Express {
    interface Request {
      appUser?: SessionPayload;
    }
  }
}

export async function hashPassword(plain: string): Promise<string> {
  return bcrypt.hash(plain, 10);
}

export async function verifyPassword(plain: string, hash: string): Promise<boolean> {
  return bcrypt.compare(plain, hash);
}

export function signSessionToken(payload: SessionPayload): string {
  return jwt.sign(payload, jwtSecret(), { expiresIn: SESSION_TTL });
}

export function verifySessionToken(token: string): SessionPayload | null {
  try {
    const decoded = jwt.verify(token, jwtSecret());
    if (typeof decoded === "string") return null;
    const { email, role } = decoded as Record<string, unknown>;
    if (typeof email !== "string" || (role !== "admin" && role !== "user")) return null;
    return { email, role };
  } catch {
    return null;
  }
}

export function requireAuth(req: Request, res: Response, next: NextFunction): void {
  const token = req.cookies?.[SESSION_COOKIE_NAME] as string | undefined;
  const session = token ? verifySessionToken(token) : null;
  if (!session) {
    res.status(401).json({ error: "not_authenticated" });
    return;
  }
  req.appUser = session;
  next();
}

export function requireAdmin(req: Request, res: Response, next: NextFunction): void {
  if (req.appUser?.role !== "admin") {
    res.status(403).json({ error: "admin_only" });
    return;
  }
  next();
}
