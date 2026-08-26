import { Router } from 'express';
import { eq, and } from 'drizzle-orm';
import { db } from '../db/client.js';
import { integrations } from '../db/schema.js';
import { getOrgId } from '../middleware/tenant.js';
import { requireRole } from '../middleware/roles.js';

const AVAILABLE_INTEGRATIONS = [
  { type: 'google_drive', name: 'Google Drive', description: 'Sync documents to your Google Drive', status: 'coming_soon' },
  { type: 'quickbooks', name: 'QuickBooks Online', description: 'Generate invoices in QuickBooks', status: 'coming_soon' },
];

const router = Router();

/** GET /api/integrations — List available + connected integrations. */
router.get('/', requireRole('owner', 'admin'), async (req, res) => {
  const orgId = getOrgId(req);
  const connected = await db.select()
    .from(integrations)
    .where(eq(integrations.orgId, orgId));

  const connectedMap = new Map(connected.map(c => [c.type, c]));

  const result = AVAILABLE_INTEGRATIONS.map(ai => ({
    ...ai,
    enabled: connectedMap.get(ai.type)?.enabled ?? false,
    connectionStatus: connectedMap.get(ai.type)?.status ?? 'disconnected',
    lastSyncAt: connectedMap.get(ai.type)?.lastSyncAt ?? null,
  }));

  res.json(result);
});

/** POST /api/integrations/:type/connect — Start connection flow. */
router.post('/:type/connect', requireRole('owner', 'admin'), async (req, res) => {
  const { type } = req.params;
  const known = AVAILABLE_INTEGRATIONS.find(a => a.type === type);
  if (!known) return res.status(404).json({ error: 'Unknown integration type' });

  // TODO: Implement OAuth flows per integration type
  // For now, just create or update the integration record
  await db.insert(integrations)
    .values({
      orgId: getOrgId(req),
      type,
      enabled: true,
      status: 'connected',
      config: req.body.config || {},
    })
    .onConflictDoUpdate({
      target: [integrations.orgId, integrations.type],
      set: {
        enabled: true,
        status: 'connected',
        config: req.body.config || {},
        updatedAt: new Date(),
      },
    });

  res.json({ ok: true, type });
});

/** POST /api/integrations/:type/disconnect — Remove connection. */
router.post('/:type/disconnect', requireRole('owner', 'admin'), async (req, res) => {
  await db.update(integrations)
    .set({ enabled: false, status: 'disconnected', credentials: null, updatedAt: new Date() })
    .where(and(eq(integrations.orgId, getOrgId(req)), eq(integrations.type, req.params.type)));

  res.json({ ok: true });
});

/** GET /api/integrations/:type/status — Connection health. */
router.get('/:type/status', async (req, res) => {
  const [integration] = await db.select()
    .from(integrations)
    .where(and(eq(integrations.orgId, getOrgId(req)), eq(integrations.type, req.params.type)))
    .limit(1);

  if (!integration) {
    return res.json({ connected: false, status: 'disconnected' });
  }
  res.json({
    connected: integration.enabled && integration.status === 'connected',
    status: integration.status,
    lastSyncAt: integration.lastSyncAt,
    error: integration.errorMessage,
  });
});

export default router;
