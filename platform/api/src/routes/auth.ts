import { Router, type Request, type Response } from 'express';
import { eq } from 'drizzle-orm';
import crypto from 'crypto';
import { db } from '../db/client.js';
import { organizations, users, invitations, overtimeRules } from '../db/schema.js';
import { hashPassword, verifyPassword } from '../utils/password.js';
import { signJwt } from '../utils/jwt.js';
import { signupSchema, loginSchema, acceptInviteSchema } from '../utils/validation.js';
import { AppError } from '../utils/errors.js';
import { authMiddleware } from '../middleware/auth.js';
import { config } from '../config.js';

const router = Router();
const SESSION_COOKIE = 'session';
const COOKIE_MAX_AGE = 7 * 24 * 60 * 60 * 1000; // 7 days

function setSessionCookie(res: Response, token: string) {
  res.cookie(SESSION_COOKIE, token, {
    httpOnly: true,
    secure: config.isProd,
    sameSite: 'lax',
    maxAge: COOKIE_MAX_AGE,
    path: '/',
  });
}

/**
 * POST /api/auth/signup
 * Creates a new organization + owner user.
 */
router.post('/signup', async (req: Request, res: Response) => {
  const parsed = signupSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ error: parsed.error.issues[0].message });
  }

  const { email, password, name, orgName } = parsed.data;

  // Check email uniqueness
  const existing = await db.select({ id: users.id })
    .from(users)
    .where(eq(users.email, email.toLowerCase()))
    .limit(1);
  if (existing.length > 0) {
    return res.status(409).json({ error: 'Email already in use' });
  }

  const passwordHash = await hashPassword(password);

  // Create org + owner in a transaction
  const result = await db.transaction(async (tx) => {
    const [org] = await tx.insert(organizations).values({
      name: orgName,
    }).returning();

    const [user] = await tx.insert(users).values({
      orgId: org.id,
      email: email.toLowerCase(),
      passwordHash,
      name,
      role: 'owner',
    }).returning();

    // Create default overtime rules for the org
    await tx.insert(overtimeRules).values({
      orgId: org.id,
    });

    return { org, user };
  });

  const token = signJwt({
    userId: result.user.id,
    orgId: result.org.id,
    role: result.user.role,
  });

  setSessionCookie(res, token);

  return res.status(201).json({
    user: {
      id: result.user.id,
      email: result.user.email,
      name: result.user.name,
      role: result.user.role,
    },
    org: {
      id: result.org.id,
      name: result.org.name,
    },
  });
});

/**
 * POST /api/auth/login
 */
router.post('/login', async (req: Request, res: Response) => {
  const parsed = loginSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ error: parsed.error.issues[0].message });
  }

  const { email, password } = parsed.data;

  const [user] = await db.select()
    .from(users)
    .where(eq(users.email, email.toLowerCase()))
    .limit(1);

  if (!user) {
    return res.status(401).json({ error: 'Invalid email or password' });
  }

  const valid = await verifyPassword(password, user.passwordHash);
  if (!valid) {
    return res.status(401).json({ error: 'Invalid email or password' });
  }

  // Update last login
  await db.update(users)
    .set({ lastLoginAt: new Date() })
    .where(eq(users.id, user.id));

  const token = signJwt({
    userId: user.id,
    orgId: user.orgId,
    role: user.role,
  });

  setSessionCookie(res, token);

  return res.json({
    user: {
      id: user.id,
      email: user.email,
      name: user.name,
      role: user.role,
    },
  });
});

/**
 * POST /api/auth/logout
 */
router.post('/logout', (_req: Request, res: Response) => {
  res.clearCookie(SESSION_COOKIE, { path: '/' });
  return res.json({ ok: true });
});

/**
 * POST /api/auth/refresh
 * Re-issues JWT from current cookie (extends session).
 */
router.post('/refresh', authMiddleware, async (req: Request, res: Response) => {
  const [user] = await db.select()
    .from(users)
    .where(eq(users.id, req.user!.userId))
    .limit(1);

  if (!user) {
    res.clearCookie(SESSION_COOKIE, { path: '/' });
    return res.status(401).json({ error: 'User not found' });
  }

  const token = signJwt({
    userId: user.id,
    orgId: user.orgId,
    role: user.role,
  });

  setSessionCookie(res, token);
  return res.json({ ok: true });
});

/**
 * GET /api/auth/me
 * Returns current user + org info.
 */
router.get('/me', authMiddleware, async (req: Request, res: Response) => {
  const [user] = await db.select({
    id: users.id,
    email: users.email,
    name: users.name,
    role: users.role,
    orgId: users.orgId,
  })
    .from(users)
    .where(eq(users.id, req.user!.userId))
    .limit(1);

  if (!user) {
    return res.status(404).json({ error: 'User not found' });
  }

  const [org] = await db.select({
    id: organizations.id,
    name: organizations.name,
    timezone: organizations.timezone,
    opDayCutoffHour: organizations.opDayCutoffHour,
  })
    .from(organizations)
    .where(eq(organizations.id, user.orgId))
    .limit(1);

  return res.json({ user, org });
});

/**
 * POST /api/auth/accept-invite
 * Accepts an email invitation and creates a user account.
 */
router.post('/accept-invite', async (req: Request, res: Response) => {
  const parsed = acceptInviteSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ error: parsed.error.issues[0].message });
  }

  const { token, name, password } = parsed.data;
  const tokenHash = crypto.createHash('sha256').update(token).digest('hex');

  const [invite] = await db.select()
    .from(invitations)
    .where(eq(invitations.tokenHash, tokenHash))
    .limit(1);

  if (!invite) {
    return res.status(404).json({ error: 'Invalid invitation' });
  }
  if (invite.acceptedAt) {
    return res.status(409).json({ error: 'Invitation already accepted' });
  }
  if (invite.expiresAt < new Date()) {
    return res.status(410).json({ error: 'Invitation expired' });
  }

  // Check email not already taken
  const existing = await db.select({ id: users.id })
    .from(users)
    .where(eq(users.email, invite.email.toLowerCase()))
    .limit(1);
  if (existing.length > 0) {
    return res.status(409).json({ error: 'Email already in use' });
  }

  const passwordHash = await hashPassword(password);

  const result = await db.transaction(async (tx) => {
    const [user] = await tx.insert(users).values({
      orgId: invite.orgId,
      email: invite.email.toLowerCase(),
      passwordHash,
      name,
      role: invite.role,
      invitedBy: invite.invitedBy,
      invitedAt: invite.createdAt,
    }).returning();

    await tx.update(invitations)
      .set({ acceptedAt: new Date() })
      .where(eq(invitations.id, invite.id));

    // Auto-create linked employee record for crew/foreman users
    // so they appear in crew chief dropdowns and sign-in sheets
    if (invite.role === 'crew' || invite.role === 'foreman') {
      const { employees } = await import('../db/schema.js');
      await tx.insert(employees).values({
        orgId: invite.orgId,
        userId: user.id,
        name: user.name,
      }).onConflictDoNothing();
    }

    return user;
  });

  const jwt = signJwt({
    userId: result.id,
    orgId: result.orgId,
    role: result.role,
  });

  setSessionCookie(res, jwt);

  return res.status(201).json({
    user: {
      id: result.id,
      email: result.email,
      name: result.name,
      role: result.role,
    },
  });
});

export default router;
