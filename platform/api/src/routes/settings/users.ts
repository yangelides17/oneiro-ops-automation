import { Router } from 'express';
import crypto from 'crypto';
import { eq, and } from 'drizzle-orm';
import { db } from '../../db/client.js';
import { listOrgUsers, updateUserRole } from '../../db/queries/users.js';
import { invitations } from '../../db/schema.js';
import { getOrgId, getUserId } from '../../middleware/tenant.js';
import { requireRole } from '../../middleware/roles.js';
import { inviteSchema } from '../../utils/validation.js';
import { resendEmail } from '../../integrations/email/resend.js';
import { config } from '../../config.js';

const router = Router();
const INVITE_EXPIRY_DAYS = 7;

router.get('/', requireRole('owner', 'admin'), async (req, res) => {
  res.json(await listOrgUsers(db, getOrgId(req)));
});

router.patch('/:id/role', requireRole('owner'), async (req, res) => {
  const { role } = req.body;
  if (!role) return res.status(400).json({ error: 'role required' });
  // M-28: Validate against allowed roles
  const ALLOWED_ROLES = ['owner', 'admin', 'foreman', 'crew'];
  if (!ALLOWED_ROLES.includes(role)) {
    return res.status(400).json({ error: `Invalid role. Must be one of: ${ALLOWED_ROLES.join(', ')}` });
  }
  const row = await updateUserRole(db, getOrgId(req), req.params.id, role);
  if (!row) return res.status(404).json({ error: 'Not found' });
  res.json(row);
});

router.post('/invite', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = inviteSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });

  const { email, role } = parsed.data;
  const token = crypto.randomBytes(32).toString('hex');
  const tokenHash = crypto.createHash('sha256').update(token).digest('hex');
  const expiresAt = new Date(Date.now() + INVITE_EXPIRY_DAYS * 24 * 60 * 60 * 1000);

  await db.insert(invitations).values({
    orgId: getOrgId(req),
    email: email.toLowerCase(),
    role,
    invitedBy: getUserId(req),
    tokenHash,
    expiresAt,
  });

  // Send invite email
  const inviteUrl = `${config.isProd ? 'https://app.oneiro.com' : 'http://localhost:5173'}/accept-invite?token=${token}`;
  try {
    await resendEmail.send(getOrgId(req), {
      to: email,
      subject: 'You\'ve been invited to Oneiro',
      html: `<p>You've been invited to join an organization on Oneiro.</p>
             <p><a href="${inviteUrl}">Accept Invitation</a></p>
             <p>This link expires in ${INVITE_EXPIRY_DAYS} days.</p>`,
    });
  } catch (e) {
    // Email failure shouldn't block invite creation
    console.error('Failed to send invite email:', e);
  }

  res.status(201).json({ ok: true, email });
});

// M-29: Delete a user from the org (owner only)
router.delete('/:id', requireRole('owner'), async (req, res) => {
  const { users } = await import('../../db/schema.js');
  const orgId = getOrgId(req);

  // Prevent self-deletion
  if (req.params.id === getUserId(req)) {
    return res.status(400).json({ error: 'Cannot remove yourself' });
  }

  const [deleted] = await db.delete(users)
    .where(and(eq(users.id, req.params.id), eq(users.orgId, orgId)))
    .returning({ id: users.id });

  if (!deleted) return res.status(404).json({ error: 'Not found' });
  res.json({ ok: true });
});

// M-29: Cancel a pending invitation
router.delete('/invite/:id', requireRole('owner', 'admin'), async (req, res) => {
  const orgId = getOrgId(req);
  const [deleted] = await db.delete(invitations)
    .where(and(eq(invitations.id, req.params.id), eq(invitations.orgId, orgId)))
    .returning({ id: invitations.id });

  if (!deleted) return res.status(404).json({ error: 'Not found' });
  res.json({ ok: true });
});

export default router;
