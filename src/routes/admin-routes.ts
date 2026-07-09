import { Router } from "express";
import {
  listUsers,
  getUser,
  upsertUser,
  deleteUser,
  getResetEmailSender,
  setResetEmailSender,
  listAppErrors,
  AppRole,
} from "../auth/app-store";
import { hashPassword, requireAuth, requireAdmin } from "../auth/app-auth";

function isValidRole(role: unknown): role is AppRole {
  return role === "admin" || role === "user";
}

async function countActiveAdmins(): Promise<number> {
  const users = await listUsers();
  return users.filter((u) => u.role === "admin" && u.isActive).length;
}

export function createAdminRoutes(getHealthSnapshot: () => Record<string, unknown>): Router {
  const router = Router();
  router.use(requireAuth, requireAdmin);

  router.get("/users", async (_req, res) => {
    const users = await listUsers();
    res.json(users.map(({ email, role, createdAt, isActive }) => ({ email, role, createdAt, isActive })));
  });

  router.post("/users", async (req, res) => {
    const { email, role, password } = req.body ?? {};
    if (typeof email !== "string" || !email.includes("@") || typeof password !== "string" || password.length < 8) {
      res.status(400).json({ error: "email and a password of at least 8 characters are required" });
      return;
    }
    if (!isValidRole(role)) {
      res.status(400).json({ error: "role must be 'admin' or 'user'" });
      return;
    }

    const existing = await getUser(email);
    if (existing) {
      res.status(409).json({ error: "user_already_exists" });
      return;
    }

    await upsertUser({
      email,
      passwordHash: await hashPassword(password),
      role,
      createdAt: Date.now(),
      isActive: true,
    });
    res.status(201).json({ ok: true });
  });

  router.patch("/users/:email", async (req, res) => {
    const targetEmail = req.params.email;
    const user = await getUser(targetEmail);
    if (!user) {
      res.status(404).json({ error: "not_found" });
      return;
    }

    const { role, isActive } = req.body ?? {};
    const nextRole: AppRole = role !== undefined ? role : user.role;
    const nextActive = isActive !== undefined ? Boolean(isActive) : user.isActive;

    if (role !== undefined && !isValidRole(role)) {
      res.status(400).json({ error: "role must be 'admin' or 'user'" });
      return;
    }

    const demotingOrDeactivatingLastAdmin =
      user.role === "admin" &&
      user.isActive &&
      (nextRole !== "admin" || !nextActive) &&
      (await countActiveAdmins()) <= 1;

    if (demotingOrDeactivatingLastAdmin) {
      res.status(400).json({ error: "cannot_remove_last_admin" });
      return;
    }

    await upsertUser({ ...user, role: nextRole, isActive: nextActive });
    res.json({ ok: true });
  });

  router.delete("/users/:email", async (req, res) => {
    const targetEmail = req.params.email;
    const user = await getUser(targetEmail);
    if (!user) {
      res.status(404).json({ error: "not_found" });
      return;
    }

    if (user.role === "admin" && user.isActive && (await countActiveAdmins()) <= 1) {
      res.status(400).json({ error: "cannot_remove_last_admin" });
      return;
    }

    await deleteUser(targetEmail);
    res.json({ ok: true });
  });

  router.get("/settings", async (_req, res) => {
    res.json({ resetEmailSender: await getResetEmailSender() });
  });

  router.put("/settings", async (req, res) => {
    const { resetEmailSender } = req.body ?? {};
    if (typeof resetEmailSender !== "string" || !resetEmailSender.includes("@")) {
      res.status(400).json({ error: "resetEmailSender must be a valid email address" });
      return;
    }
    await setResetEmailSender(resetEmailSender);
    res.json({ ok: true });
  });

  router.get("/logs", async (_req, res) => {
    const [errors, health] = await Promise.all([
      listAppErrors(100),
      Promise.resolve(getHealthSnapshot()),
    ]);
    res.json({ errors, health });
  });

  return router;
}
