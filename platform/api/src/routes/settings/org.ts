import { Router } from 'express';
import { db } from '../../db/client.js';
import { getOrg, updateOrg } from '../../db/queries/organizations.js';
import { getOrgId } from '../../middleware/tenant.js';
import { requireRole } from '../../middleware/roles.js';
import { updateOrgSchema } from '../../utils/validation.js';

const router = Router();

router.get('/', async (req, res) => {
  const org = await getOrg(db, getOrgId(req));
  if (!org) return res.status(404).json({ error: 'Organization not found' });
  res.json(org);
});

router.patch('/', requireRole('owner'), async (req, res) => {
  const parsed = updateOrgSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const updated = await updateOrg(db, getOrgId(req), parsed.data);
  res.json(updated);
});

export default router;
