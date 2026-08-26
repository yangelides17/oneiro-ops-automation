import { Router } from 'express';
import multer from 'multer';
import { db } from '../db/client.js';
import { listPhotos, getPhoto, createPhoto, deletePhoto } from '../db/queries/photos.js';
import { getWorkOrder } from '../db/queries/workOrders.js';
import { getOrgId } from '../middleware/tenant.js';
import { requireRole } from '../middleware/roles.js';
import { r2Storage } from '../integrations/storage/r2.js';

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });
const router = Router();

/** POST /api/wos/:woId/photos — Upload a photo. */
router.post('/wos/:woId/photos', requireRole('owner', 'admin', 'foreman', 'crew'), upload.single('photo'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file attached' });

  const orgId = getOrgId(req);
  const woId = req.params.woId;

  const wo = await getWorkOrder(db, orgId, woId);
  if (!wo) return res.status(404).json({ error: 'Work order not found' });

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

  // Include old-app field names the frontend's photoUploadQueue expects
  const url = await r2Storage.getSignedUrl(orgId, storageKey, 3600).catch(() => null);
  res.status(201).json({ ...photo, fileId: photo.id, file_url: url });
});

/** GET /api/wos/:woId/photos — List photos for a WO. */
router.get('/wos/:woId/photos', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const orgId = getOrgId(req);
  const items = await listPhotos(db, orgId, req.params.woId);

  const withUrls = await Promise.all(items.map(async (p) => ({
    ...p,
    url: await r2Storage.getSignedUrl(orgId, p.storageKey, 3600).catch(() => null),
    thumbnailUrl: p.thumbnailKey
      ? await r2Storage.getSignedUrl(orgId, p.thumbnailKey, 3600).catch(() => null)
      : null,
  })));

  res.json({ photos: withUrls });
});

/** GET /api/photos/:id/content — Full-res photo content. */
router.get('/:id/content', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const orgId = getOrgId(req);
  const photo = await getPhoto(db, orgId, req.params.id);
  if (!photo) return res.status(404).json({ error: 'Not found' });

  try {
    const data = await r2Storage.download(orgId, photo.storageKey);
    res.setHeader('Content-Type', photo.mimeType);
    res.setHeader('Content-Length', data.length);
    res.setHeader('Cache-Control', 'private, max-age=3600');
    res.send(data);
  } catch {
    res.status(500).json({ error: 'Failed to retrieve photo' });
  }
});

/** DELETE /api/photos/:id — Delete a photo. */
router.delete('/:id', requireRole('owner', 'admin', 'foreman'), async (req, res) => {
  const orgId = getOrgId(req);
  const photo = await deletePhoto(db, orgId, req.params.id);
  if (!photo) return res.status(404).json({ error: 'Not found' });

  r2Storage.delete(orgId, photo.storageKey).catch(() => {});
  if (photo.thumbnailKey) {
    r2Storage.delete(orgId, photo.thumbnailKey).catch(() => {});
  }

  res.json({ ok: true });
});

export default router;
