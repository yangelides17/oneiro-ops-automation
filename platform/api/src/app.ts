import 'express-async-errors';
import express from 'express';
import cors from 'cors';
import cookieParser from 'cookie-parser';
import path from 'path';
import { fileURLToPath } from 'url';
import { config } from './config.js';
import { errorHandler } from './utils/errors.js';
import { authMiddleware } from './middleware/auth.js';
import { tenantMiddleware, getOrgId } from './middleware/tenant.js';
import { requireRole } from './middleware/roles.js';
import { db } from './db/client.js';
import multer from 'multer';

// Routes
import authRoutes from './routes/auth.js';
import settingsRoutes from './routes/settings/index.js';
import workOrderRoutes from './routes/workOrders.js';
import markingItemRoutes from './routes/markingItems.js';
import signinRoutes from './routes/signin.js';
import fieldReportRoutes from './routes/fieldReports.js';
import documentRoutes from './routes/documents.js';
import photoRoutes from './routes/photos.js';
import toolRoutes from './routes/tools.js';
import dashboardRoutes from './routes/dashboards.js';
import geocodingRoutes from './routes/geocoding.js';
import integrationRoutes from './routes/integrations.js';
import fileRoutes from './routes/files.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export function createApp() {
  const app = express();

  // ─── Global Middleware ──────────────────────────────────────
  app.use(cors({
    origin: config.isProd ? undefined : 'http://localhost:5173',
    credentials: true,
  }));
  app.use(express.json({ limit: '10mb' }));
  app.use(cookieParser());

  // ─── Health Check ───────────────────────────────────────────
  app.get('/api/health', (_req, res) => {
    res.json({ ok: true, env: config.nodeEnv });
  });

  // ─── Auth Routes (no auth required) ─────────────────────────
  app.use('/api/auth', authRoutes);

  // ─── Authenticated Routes ───────────────────────────────────
  // All routes below require authentication + tenant context
  app.use('/api', authMiddleware, tenantMiddleware);

  // Settings routes
  app.use('/api/settings', settingsRoutes);

  // Core operational routes
  app.use('/api/wos', workOrderRoutes);
  app.use('/api', markingItemRoutes);        // Mounts /api/wos/:woId/markings + /api/markings/:id
  app.use('/api/signin', signinRoutes);
  app.use('/api/field-reports', fieldReportRoutes);
  app.use('/api/documents', documentRoutes);
  app.use('/api', photoRoutes);              // Mounts /api/wos/:woId/photos + /api/photos/:id
  app.use('/api/tools', toolRoutes);
  app.use('/api', dashboardRoutes);          // Mounts /api/dashboard, /api/revenue, /api/production, /api/pending-counts
  app.use('/api', geocodingRoutes);          // Mounts /api/geocode, /api/reverse-geocode
  app.use('/api/integrations', integrationRoutes);
  app.use('/api', fileRoutes);              // Mounts /api/wos/:woId/files + /api/files

  // ─── Additional endpoints the React frontend expects ─────────

  // /api/qb/status — QB is deferred, return disconnected stub
  app.get('/api/qb/status', (_req, res) => {
    res.json({ connected: false, reason: 'not_configured' });
  });

  // /api/employees — all active employees (for sign-in sheets, payroll)
  app.get('/api/employees', async (req, res) => {
    const { listEmployees } = await import('./db/queries/settings.js');
    const emps = await listEmployees(db, getOrgId(req));
    res.json({ employees: emps.map((e: any) => ({ name: e.name })) });
  });

  // /api/crew-chiefs — employees with linked user accounts (crew chief dropdown)
  app.get('/api/crew-chiefs', async (req, res) => {
    const { employees: empTable } = await import('./db/schema.js');
    const { eq: eqOp, and: andOp, isNotNull } = await import('drizzle-orm');
    const orgId = getOrgId(req);
    const chiefs = await db.select({ id: empTable.id, name: empTable.name, userId: empTable.userId })
      .from(empTable)
      .where(andOp(eqOp(empTable.orgId, orgId), eqOp(empTable.isActive, true), isNotNull(empTable.userId)))
      .orderBy(empTable.name);
    res.json({ employees: chiefs.map((e: any) => ({ name: e.name, userId: e.userId })) });
  });

  // /api/qb/invoice/:woId — QB deferred, stub
  app.post('/api/qb/invoice/:woId', (_req, res) => {
    res.status(503).json({ error: 'QuickBooks integration not yet configured' });
  });

  // /api/wo-markings/:woId — FieldReport may still use this path
  app.get('/api/wo-markings/:woId', async (req, res) => {
    const { listMarkingItems } = await import('./db/queries/markingItems.js');
    const { getWorkOrderByNumber } = await import('./db/queries/workOrders.js');
    const orgId = getOrgId(req);

    let woId = req.params.woId;
    if (!woId.match(/^[0-9a-f]{8}-/i)) {
      const wo = await getWorkOrderByNumber(db, orgId, woId);
      if (!wo) return res.status(404).json({ error: 'Work order not found' });
      woId = wo.id;
    }

    const items = await listMarkingItems(db, orgId, woId);
    res.json({ items });
  });

  // POST /api/markings — MarkingFormModal posts here with woId in body
  app.post('/api/markings', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
    const { createMarkingItem } = await import('./db/queries/markingItems.js');
    const { getWorkOrderByNumber } = await import('./db/queries/workOrders.js');
    const { enforceCreateDefaults, validateGridCategory } = await import('./services/markingItemLogic.js');
    const orgId = getOrgId(req);
    const { woId: woIdParam, quantity, ...rest } = req.body;
    if (!woIdParam) return res.status(400).json({ error: 'woId required' });

    let woId = woIdParam;
    if (!woId.match(/^[0-9a-f]{8}-/i)) {
      const wo = await getWorkOrderByNumber(db, orgId, woId);
      if (!wo) return res.status(404).json({ error: 'Work order not found' });
      woId = wo.id;
    }

    // Enforce same business rules as the canonical route
    const gridErr = validateGridCategory(rest.category, rest.intersection, rest.direction);
    if (gridErr) return res.status(400).json({ error: gridErr });

    const defaults = enforceCreateDefaults({ ...rest, quantity });
    const item = await createMarkingItem(db, orgId, woId, {
      ...defaults,
      quantity: defaults.quantity != null ? String(defaults.quantity) : undefined,
    });
    res.status(201).json({ item });
  });

  // /api/wo-photos/:woId — Photo list (frontend expects old field names)
  app.get('/api/wo-photos/:woId', async (req, res) => {
    const { listPhotos } = await import('./db/queries/photos.js');
    const { getWorkOrderByNumber } = await import('./db/queries/workOrders.js');
    const { r2Storage } = await import('./integrations/storage/r2.js');
    const orgId = getOrgId(req);

    let woId = req.params.woId;
    if (!woId.match(/^[0-9a-f]{8}-/i)) {
      const wo = await getWorkOrderByNumber(db, orgId, woId);
      if (!wo) return res.status(404).json({ error: 'Work order not found' });
      woId = wo.id;
    }

    const photos = await listPhotos(db, orgId, woId);
    // Map to old-app field names the frontend expects
    const mapped = await Promise.all((photos || []).map(async (p: any) => ({
      fileId: p.id,
      name: p.filename,
      mime: p.mimeType,
      url: await r2Storage.getSignedUrl(orgId, p.storageKey, 3600).catch(() => null),
      thumbnail_b64: null, // R2 doesn't store thumbnails as base64
      created_at: p.createdAt,
    })));
    res.json({ photos: mapped });
  });

  // /api/wo-photos/:fileId/content — Photo content by DB id
  app.get('/api/wo-photos/:fileId/content', async (req, res) => {
    const { getPhoto } = await import('./db/queries/photos.js');
    const { r2Storage } = await import('./integrations/storage/r2.js');
    const orgId = getOrgId(req);
    const photo = await getPhoto(db, orgId, req.params.fileId);
    if (!photo) return res.status(404).json({ error: 'Photo not found' });
    try {
      const data = await r2Storage.download(orgId, photo.storageKey);
      res.setHeader('Content-Type', photo.mimeType);
      res.setHeader('Content-Length', data.length);
      res.send(data);
    } catch {
      res.status(500).json({ error: 'Failed to retrieve photo' });
    }
  });

  // DELETE /api/wo-photos/:fileId — Delete a photo by DB id
  app.delete('/api/wo-photos/:fileId', requireRole('owner', 'admin', 'foreman'), async (req, res) => {
    const { deletePhoto } = await import('./db/queries/photos.js');
    const { r2Storage } = await import('./integrations/storage/r2.js');
    const orgId = getOrgId(req);
    const photo = await deletePhoto(db, orgId, req.params.fileId);
    if (!photo) return res.status(404).json({ error: 'Photo not found' });
    r2Storage.delete(orgId, photo.storageKey).catch(() => {});
    res.json({ ok: true });
  });

  // /api/upload-photo — compat route for photoUploadQueue.js
  const photoUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });
  app.post('/api/upload-photo', requireRole('owner', 'admin', 'foreman', 'crew'), photoUpload.single('photo'), async (req: any, res) => {
    if (!req.file) return res.status(400).json({ error: 'No file attached' });
    const orgId = getOrgId(req);
    const woId = req.body.woId;
    if (!woId) return res.status(400).json({ error: 'woId required' });

    const { getWorkOrder } = await import('./db/queries/workOrders.js');
    const wo = await getWorkOrder(db, orgId, woId);
    if (!wo) return res.status(404).json({ error: 'Work order not found' });

    const { r2Storage } = await import('./integrations/storage/r2.js');
    const { createPhoto } = await import('./db/queries/photos.js');

    const path = `photos/${woId}/${Date.now()}_${req.file.originalname}`;
    const storageKey = await r2Storage.upload(orgId, path, req.file.buffer, req.file.mimetype);

    const photo = await createPhoto(db, orgId, {
      orgId,
      woId,
      storageKey,
      filename: req.file.originalname,
      mimeType: req.file.mimetype,
      sizeBytes: req.file.size,
      latitude: req.body.latitude ? String(req.body.latitude) : undefined,
      longitude: req.body.longitude ? String(req.body.longitude) : undefined,
      address: req.body.address || undefined,
    });

    const url = await r2Storage.getSignedUrl(orgId, storageKey, 3600).catch(() => null);
    res.json({ success: true, fileId: photo.id, file_id: photo.id, file_url: url });
  });

  // /api/signin/queue/parse-upload — parse uploaded sign-in PDF via Claude Vision
  app.post('/api/signin/queue/parse-upload', async (_req, res) => {
    // Requires Python worker + Claude Vision integration
    res.json({ parsed: null, message: 'Sign-in PDF parsing requires worker integration' });
  });

  // ─── Error Handler ──────────────────────────────────────────
  app.use(errorHandler);

  // ─── Static Files (production: serve React SPA) ─────────────
  if (config.isProd) {
    const webDist = path.resolve(__dirname, '../../web/dist');
    app.use(express.static(webDist));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(webDist, 'index.html'));
    });
  }

  return app;
}
