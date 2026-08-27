/**
 * Files API — unified file browsing across scans, documents, and photos.
 *
 * Endpoints:
 *   GET /api/wos/:woId/files       — all files for one WO (metadata + proxy URLs)
 *   GET /api/wos/:woId/scan.pdf    — stream WO scan PDF (same-origin, iframe-safe)
 *   GET /api/files                  — paginated browser with filters
 *   GET /api/files/:id/view        — stream any file by ID + type (same-origin)
 */

import { Router } from 'express';
import { eq, and } from 'drizzle-orm';
import { db } from '../db/client.js';
import { getWorkOrder, getWorkOrderByNumber } from '../db/queries/workOrders.js';
import {
  listFilesForWO, listFilesPaginated, type FileFilters,
  getContractorsWithFiles, getContractRegionsForContractor, getWosForContractRegion,
} from '../db/queries/files.js';
import { getOrgId } from '../middleware/tenant.js';
import { requireRole } from '../middleware/roles.js';
import { r2Storage } from '../integrations/storage/r2.js';
import { workOrders, documents, photos } from '../db/schema.js';

const router = Router();

// ─── Helper: resolve WO ID from UUID or number ───────────────

async function resolveWoId(orgId: string, param: string) {
  if (param.match(/^[0-9a-f]{8}-/i)) {
    return await getWorkOrder(db, orgId, param);
  }
  return await getWorkOrderByNumber(db, orgId, param);
}

// ─── Per-WO Files ─────────────────────────────────────────────

/**
 * GET /api/wos/:woId/files — All files for a work order.
 *
 * Returns metadata with same-origin proxy URLs (not signed R2 URLs)
 * so iframes can render PDFs without cross-origin blocks.
 */
router.get('/wos/:woId/files', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const orgId = getOrgId(req);
  const wo = await resolveWoId(orgId, req.params.woId);
  if (!wo) return res.status(404).json({ error: 'Work order not found' });

  const { scan, documents: docs, photos: photoRows } = await listFilesForWO(db, orgId, wo.id, wo.woNumber);

  // Use same-origin proxy URLs instead of signed R2 URLs
  const scanWithUrl = scan
    ? {
        storageKey: scan.storageKey,
        filename: scan.filename,
        mimeType: 'application/pdf',
        url: `/api/wos/${wo.id}/scan.pdf`,
      }
    : null;

  const docsWithUrls = docs
    .filter(d => d.storageKey)
    .map(d => ({
      id: d.id,
      docType: d.docType,
      filename: d.filename || `${d.docType}.pdf`,
      anchorDate: d.anchorDate,
      status: d.status,
      url: `/api/documents/${d.id}/pdf`,
    }));

  const photosWithUrls = photoRows.map(p => ({
    id: p.id,
    filename: p.filename || 'photo.jpg',
    mimeType: p.mimeType,
    sizeBytes: p.sizeBytes,
    createdAt: p.createdAt,
    url: `/api/files/${p.id}/view?type=photo`,
  }));

  res.json({ scan: scanWithUrl, documents: docsWithUrls, photos: photosWithUrls });
});

/**
 * GET /api/wos/:woId/scan.pdf — Stream the WO scan PDF.
 *
 * Same-origin endpoint so Chrome renders it in an iframe without
 * cross-origin blocks. This replaces the signed R2 URL approach.
 */
router.get('/wos/:woId/scan.pdf', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const orgId = getOrgId(req);
  const wo = await resolveWoId(orgId, req.params.woId);
  if (!wo || !wo.scanFileKey) return res.status(404).json({ error: 'Scan not found' });

  try {
    const data = await r2Storage.download(orgId, wo.scanFileKey);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Length', data.length);
    res.setHeader('Content-Disposition', `inline; filename="${wo.originalFilename || wo.woNumber + '.pdf'}"`);
    res.send(data);
  } catch {
    res.status(500).json({ error: 'Failed to retrieve scan' });
  }
});

/**
 * GET /api/files/:id/view — Stream any file by its ID.
 *
 * Query param: type=photo|document|scan
 * Same-origin proxy for embedding in iframes and img tags.
 */
router.get('/files/:id/view', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const orgId = getOrgId(req);
  const { id } = req.params;
  const type = req.query.type as string;

  let storageKey: string | null = null;
  let mimeType = 'application/octet-stream';
  let filename = 'file';

  if (type === 'photo') {
    const [photo] = await db.select({
      storageKey: photos.storageKey,
      mimeType: photos.mimeType,
      filename: photos.filename,
    })
      .from(photos)
      .where(and(eq(photos.id, id), eq(photos.orgId, orgId)))
      .limit(1);
    if (photo) {
      storageKey = photo.storageKey;
      mimeType = photo.mimeType || 'image/jpeg';
      filename = photo.filename || 'photo.jpg';
    }
  } else if (type === 'document') {
    const [doc] = await db.select({
      storageKey: documents.storageKey,
      filename: documents.filename,
    })
      .from(documents)
      .where(and(eq(documents.id, id), eq(documents.orgId, orgId)))
      .limit(1);
    if (doc?.storageKey) {
      storageKey = doc.storageKey;
      mimeType = 'application/pdf';
      filename = doc.filename || 'document.pdf';
    }
  } else if (type === 'scan') {
    const [wo] = await db.select({
      scanFileKey: workOrders.scanFileKey,
      originalFilename: workOrders.originalFilename,
      woNumber: workOrders.woNumber,
    })
      .from(workOrders)
      .where(and(eq(workOrders.id, id), eq(workOrders.orgId, orgId)))
      .limit(1);
    if (wo?.scanFileKey) {
      storageKey = wo.scanFileKey;
      mimeType = 'application/pdf';
      filename = wo.originalFilename || `${wo.woNumber}.pdf`;
    }
  }

  if (!storageKey) return res.status(404).json({ error: 'File not found' });

  try {
    const data = await r2Storage.download(orgId, storageKey);
    res.setHeader('Content-Type', mimeType);
    res.setHeader('Content-Length', data.length);
    res.setHeader('Content-Disposition', `inline; filename="${filename}"`);
    res.send(data);
  } catch {
    res.status(500).json({ error: 'Failed to retrieve file' });
  }
});

// ─── Virtual Folder Browse ────────────────────────────────────

/**
 * GET /api/files/browse — Navigate the virtual folder hierarchy.
 *
 * Query params determine the level:
 *   (none)                          → Level 0: contractors with files
 *   contractorId=xxx                → Level 1: contract/region combos
 *   contractorId&contractNum&regionCode → Level 2: WOs with file counts
 */
router.get('/files/browse', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const orgId = getOrgId(req);
  const { contractorId, contractNum, regionCode } = req.query as Record<string, string>;

  if (contractorId && contractNum && regionCode) {
    // Level 2: WOs for a contract+region
    const items = await getWosForContractRegion(db, orgId, contractNum, regionCode);
    return res.json({ level: 'workorders', items });
  }

  if (contractorId) {
    // Level 1: contract+region combos for a contractor
    const items = await getContractRegionsForContractor(db, orgId, contractorId);
    return res.json({ level: 'contracts', items });
  }

  // Level 0: contractors with files
  const items = await getContractorsWithFiles(db, orgId);
  res.json({ level: 'contractors', items });
});

// ─── Paginated File Browser ───────────────────────────────────

/**
 * GET /api/files — Paginated file browser.
 */
router.get('/files', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const orgId = getOrgId(req);
  const page = Math.max(1, parseInt(req.query.page as string, 10) || 1);
  const limit = Math.min(100, Math.max(1, parseInt(req.query.limit as string, 10) || 50));

  const filters: FileFilters = {};
  if (req.query.type) filters.type = req.query.type as FileFilters['type'];
  if (req.query.docType) filters.docType = req.query.docType as string;
  if (req.query.contractorId) filters.contractorId = req.query.contractorId as string;
  if (req.query.dateStart) filters.dateStart = req.query.dateStart as string;
  if (req.query.dateEnd) filters.dateEnd = req.query.dateEnd as string;
  if (req.query.search) filters.search = req.query.search as string;

  const { files, total } = await listFilesPaginated(db, orgId, filters, page, limit);

  // Use same-origin proxy URLs instead of signed R2 URLs
  const filesWithUrls = files.map(f => {
    let url: string;
    if (f.type === 'scan') {
      url = `/api/files/${f.id}/view?type=scan`;
    } else if (f.type === 'document') {
      url = `/api/files/${f.id}/view?type=document`;
    } else {
      url = `/api/files/${f.id}/view?type=photo`;
    }
    return { ...f, url };
  });

  res.json({ files: filesWithUrls, total, page, pageSize: limit });
});

export default router;
