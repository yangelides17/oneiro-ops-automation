/**
 * File Queries — aggregates across work_orders, documents, and photos.
 *
 * Provides a unified view of all files in the system without a new table.
 * The three file sources have different schemas, so we run separate queries
 * and merge in JS rather than attempting a SQL UNION.
 */

import { eq, and, between, sql, desc, asc, ilike } from 'drizzle-orm';
import type { Db } from '../client.js';
import { workOrders, documents, photos, contractors } from '../schema.js';

// ─── Types ────────────────────────────────────────────────────

export interface UnifiedFile {
  id: string;
  type: 'scan' | 'document' | 'photo';
  docType?: string;             // field_report, production_log, signin, certified_payroll
  filename: string;
  storageKey: string;
  mimeType: string;
  woNumber: string;
  woId: string;
  contractorName: string;
  contractorId: string;
  contractNum: string;
  regionCode: string;
  date: string;                 // anchorDate for docs, dateCompleted for scans, createdAt for photos
  status?: string;              // document status (pending, needs_review, approved, archived)
  sizeBytes?: number;
  createdAt: Date | string;
}

export interface FileFilters {
  type?: 'scan' | 'document' | 'photo';
  docType?: string;
  contractorId?: string;
  dateStart?: string;
  dateEnd?: string;
  search?: string;              // WO number search
}

// ─── Per-WO Files ─────────────────────────────────────────────

/**
 * Get all files for a specific WO: scan + documents + photos.
 * Used by the inline PDF viewer on Field Report and Map tab.
 */
export async function listFilesForWO(
  db: Db,
  orgId: string,
  woId: string,
  woNumber: string,
) {
  // 1. Scan file (from work_orders.scanFileKey)
  const [wo] = await db.select({
    scanFileKey: workOrders.scanFileKey,
    originalFilename: workOrders.originalFilename,
  })
    .from(workOrders)
    .where(and(eq(workOrders.id, woId), eq(workOrders.orgId, orgId)))
    .limit(1);

  const scan = wo?.scanFileKey
    ? { storageKey: wo.scanFileKey, filename: wo.originalFilename || `${woNumber}.pdf` }
    : null;

  // 2. Documents (where woIds array contains this WO number)
  const docs = await db.select({
    id: documents.id,
    docType: documents.docType,
    docKey: documents.docKey,
    filename: documents.filename,
    storageKey: documents.storageKey,
    anchorDate: documents.anchorDate,
    status: documents.status,
    createdAt: documents.createdAt,
  })
    .from(documents)
    .where(and(
      eq(documents.orgId, orgId),
      sql`${documents.woIds} @> ARRAY[${woNumber}]::text[]`,
    ))
    .orderBy(desc(documents.createdAt));

  // 3. Photos
  const photoRows = await db.select({
    id: photos.id,
    filename: photos.filename,
    storageKey: photos.storageKey,
    mimeType: photos.mimeType,
    sizeBytes: photos.sizeBytes,
    createdAt: photos.createdAt,
  })
    .from(photos)
    .where(and(eq(photos.orgId, orgId), eq(photos.woId, woId)))
    .orderBy(desc(photos.createdAt));

  return { scan, documents: docs, photos: photoRows };
}

// ─── Paginated File Browser ───────────────────────────────────

/**
 * Paginated query across all file types for the file browser.
 * Runs three separate queries (scans, documents, photos), merges,
 * sorts by date descending, and returns the requested page.
 */
export async function listFilesPaginated(
  db: Db,
  orgId: string,
  filters: FileFilters,
  page: number = 1,
  limit: number = 50,
): Promise<{ files: UnifiedFile[]; total: number }> {
  const results: UnifiedFile[] = [];

  const shouldInclude = (type: 'scan' | 'document' | 'photo') =>
    !filters.type || filters.type === type;

  // ── Scans ───────────────────────────────────────────────────
  if (shouldInclude('scan') && !filters.docType) {
    const scanConditions = [
      eq(workOrders.orgId, orgId),
      sql`${workOrders.scanFileKey} IS NOT NULL`,
    ];
    if (filters.contractorId) {
      scanConditions.push(eq(workOrders.contractorId, filters.contractorId));
    }
    if (filters.search) {
      scanConditions.push(ilike(workOrders.woNumber, `%${filters.search}%`));
    }
    if (filters.dateStart) {
      scanConditions.push(sql`${workOrders.createdAt}::date >= ${filters.dateStart}`);
    }
    if (filters.dateEnd) {
      scanConditions.push(sql`${workOrders.createdAt}::date <= ${filters.dateEnd}`);
    }

    const scans = await db.select({
      id: workOrders.id,
      scanFileKey: workOrders.scanFileKey,
      originalFilename: workOrders.originalFilename,
      woNumber: workOrders.woNumber,
      contractorId: workOrders.contractorId,
      contractorName: contractors.name,
      contractNum: workOrders.contractNum,
      regionCode: workOrders.regionCode,
      createdAt: workOrders.createdAt,
    })
      .from(workOrders)
      .leftJoin(contractors, eq(workOrders.contractorId, contractors.id))
      .where(and(...scanConditions));

    for (const s of scans) {
      results.push({
        id: s.id,
        type: 'scan',
        filename: s.originalFilename || `${s.woNumber}.pdf`,
        storageKey: s.scanFileKey!,
        mimeType: 'application/pdf',
        woNumber: s.woNumber,
        woId: s.id,
        contractorName: s.contractorName || '',
        contractorId: s.contractorId,
        contractNum: s.contractNum || '',
        regionCode: s.regionCode || '',
        date: s.createdAt?.toISOString?.()?.slice(0, 10) || '',
        createdAt: s.createdAt!,
      });
    }
  }

  // ── Documents ───────────────────────────────────────────────
  if (shouldInclude('document')) {
    const docConditions = [eq(documents.orgId, orgId)];
    if (filters.docType) {
      docConditions.push(eq(documents.docType, filters.docType));
    }
    if (filters.contractorId) {
      docConditions.push(eq(documents.contractorId, filters.contractorId));
    }
    if (filters.search) {
      docConditions.push(sql`EXISTS (
        SELECT 1 FROM unnest(${documents.woIds}) AS wo_num
        WHERE wo_num ILIKE ${'%' + filters.search + '%'}
      )`);
    }
    if (filters.dateStart && filters.dateEnd) {
      docConditions.push(between(documents.anchorDate, filters.dateStart, filters.dateEnd));
    } else if (filters.dateStart) {
      docConditions.push(sql`${documents.anchorDate} >= ${filters.dateStart}`);
    } else if (filters.dateEnd) {
      docConditions.push(sql`${documents.anchorDate} <= ${filters.dateEnd}`);
    }
    // Only show documents that have a storage key (actually filled)
    docConditions.push(sql`${documents.storageKey} IS NOT NULL`);

    const docs = await db.select({
      id: documents.id,
      docType: documents.docType,
      filename: documents.filename,
      storageKey: documents.storageKey,
      anchorDate: documents.anchorDate,
      status: documents.status,
      contractorId: documents.contractorId,
      contractorName: contractors.name,
      contractNum: documents.contractNum,
      regionCode: documents.regionCode,
      woIds: documents.woIds,
      createdAt: documents.createdAt,
    })
      .from(documents)
      .leftJoin(contractors, eq(documents.contractorId, contractors.id))
      .where(and(...docConditions));

    for (const d of docs) {
      results.push({
        id: d.id,
        type: 'document',
        docType: d.docType,
        filename: d.filename || `${d.docType}_${d.anchorDate || ''}.pdf`,
        storageKey: d.storageKey!,
        mimeType: 'application/pdf',
        woNumber: d.woIds?.[0] || '',
        woId: '',
        contractorName: d.contractorName || '',
        contractorId: d.contractorId || '',
        contractNum: d.contractNum || '',
        regionCode: d.regionCode || '',
        date: d.anchorDate || '',
        status: d.status,
        createdAt: d.createdAt!,
      });
    }
  }

  // ── Photos ──────────────────────────────────────────────────
  if (shouldInclude('photo') && !filters.docType) {
    const photoConditions = [eq(photos.orgId, orgId)];
    if (filters.contractorId) {
      // Join through workOrders to filter by contractor
      photoConditions.push(sql`${photos.woId} IN (
        SELECT id FROM work_orders WHERE org_id = ${orgId} AND contractor_id = ${filters.contractorId}
      )`);
    }
    if (filters.search) {
      photoConditions.push(sql`${photos.woId} IN (
        SELECT id FROM work_orders WHERE org_id = ${orgId} AND wo_number ILIKE ${'%' + filters.search + '%'}
      )`);
    }
    if (filters.dateStart) {
      photoConditions.push(sql`${photos.createdAt}::date >= ${filters.dateStart}`);
    }
    if (filters.dateEnd) {
      photoConditions.push(sql`${photos.createdAt}::date <= ${filters.dateEnd}`);
    }

    const photoRows = await db.select({
      id: photos.id,
      filename: photos.filename,
      storageKey: photos.storageKey,
      mimeType: photos.mimeType,
      sizeBytes: photos.sizeBytes,
      woId: photos.woId,
      createdAt: photos.createdAt,
    })
      .from(photos)
      .where(and(...photoConditions));

    // Batch-resolve WO numbers for photos
    const photoWoIds = [...new Set(photoRows.map(p => p.woId))];
    const woLookup = new Map<string, { woNumber: string; contractorName: string; contractorId: string; contractNum: string; regionCode: string }>();
    if (photoWoIds.length > 0) {
      const { inArray } = await import('drizzle-orm');
      const woRows = await db.select({
        id: workOrders.id,
        woNumber: workOrders.woNumber,
        contractorId: workOrders.contractorId,
        contractorName: contractors.name,
        contractNum: workOrders.contractNum,
        regionCode: workOrders.regionCode,
      })
        .from(workOrders)
        .leftJoin(contractors, eq(workOrders.contractorId, contractors.id))
        .where(inArray(workOrders.id, photoWoIds));
      for (const w of woRows) {
        woLookup.set(w.id, {
          woNumber: w.woNumber,
          contractorName: w.contractorName || '',
          contractorId: w.contractorId,
          contractNum: w.contractNum || '',
          regionCode: w.regionCode || '',
        });
      }
    }

    for (const p of photoRows) {
      const wo = woLookup.get(p.woId) || { woNumber: '', contractorName: '', contractorId: '', contractNum: '', regionCode: '' };
      results.push({
        id: p.id,
        type: 'photo',
        filename: p.filename || 'photo.jpg',
        storageKey: p.storageKey,
        mimeType: p.mimeType || 'image/jpeg',
        woNumber: wo.woNumber,
        woId: p.woId,
        contractorName: wo.contractorName,
        contractorId: wo.contractorId,
        contractNum: wo.contractNum,
        regionCode: wo.regionCode,
        date: p.createdAt?.toISOString?.()?.slice(0, 10) || '',
        sizeBytes: p.sizeBytes ?? undefined,
        createdAt: p.createdAt!,
      });
    }
  }

  // ── Sort and paginate ───────────────────────────────────────
  results.sort((a, b) => {
    const da = new Date(a.createdAt).getTime();
    const db_ = new Date(b.createdAt).getTime();
    return db_ - da; // newest first
  });

  const total = results.length;
  const start = (page - 1) * limit;
  const paged = results.slice(start, start + limit);

  return { files: paged, total };
}

// ─── Virtual Folder Hierarchy ─────────────────────────────────

/**
 * Level 0: Contractors that have at least one file (scan, doc, or photo).
 * Returns distinct contractors with a count of WOs that have files.
 */
export async function getContractorsWithFiles(db: Db, orgId: string) {
  // WOs with scans give us contractor_id directly
  const rows = await db.select({
    contractorId: workOrders.contractorId,
    contractorName: contractors.name,
    woCount: sql<number>`count(distinct ${workOrders.id})`.as('woCount'),
  })
    .from(workOrders)
    .leftJoin(contractors, eq(workOrders.contractorId, contractors.id))
    .where(and(
      eq(workOrders.orgId, orgId),
      sql`(${workOrders.scanFileKey} IS NOT NULL
        OR EXISTS (SELECT 1 FROM photos p WHERE p.wo_id = ${workOrders.id})
        OR EXISTS (SELECT 1 FROM documents d WHERE d.org_id = ${orgId} AND d.wo_ids @> ARRAY[${workOrders.woNumber}]::text[] AND d.storage_key IS NOT NULL)
      )`,
    ))
    .groupBy(workOrders.contractorId, contractors.name);

  return rows.map(r => ({
    id: r.contractorId,
    label: r.contractorName || 'Unknown',
    count: Number(r.woCount),
  }));
}

/**
 * Level 1: Contract+region combos for a contractor that have files.
 */
export async function getContractRegionsForContractor(
  db: Db, orgId: string, contractorId: string,
) {
  const rows = await db.select({
    contractNum: workOrders.contractNum,
    regionCode: workOrders.regionCode,
    woCount: sql<number>`count(distinct ${workOrders.id})`.as('woCount'),
  })
    .from(workOrders)
    .where(and(
      eq(workOrders.orgId, orgId),
      eq(workOrders.contractorId, contractorId),
      sql`(${workOrders.scanFileKey} IS NOT NULL
        OR EXISTS (SELECT 1 FROM photos p WHERE p.wo_id = ${workOrders.id})
        OR EXISTS (SELECT 1 FROM documents d WHERE d.org_id = ${orgId} AND d.wo_ids @> ARRAY[${workOrders.woNumber}]::text[] AND d.storage_key IS NOT NULL)
      )`,
    ))
    .groupBy(workOrders.contractNum, workOrders.regionCode);

  const BOROUGH_NAMES: Record<string, string> = {
    M: 'Manhattan', BX: 'Bronx', BK: 'Brooklyn', QU: 'Queens', SI: 'Staten Island',
  };

  return rows.map(r => ({
    contractNum: r.contractNum || '',
    regionCode: r.regionCode || '',
    label: r.contractNum || 'Unknown',
    sublabel: BOROUGH_NAMES[r.regionCode || ''] || r.regionCode || '',
    count: Number(r.woCount),
  }));
}

/**
 * Level 2: WOs for a contract+region with file counts.
 */
export async function getWosForContractRegion(
  db: Db, orgId: string, contractNum: string, regionCode: string,
) {
  const rows = await db.select({
    id: workOrders.id,
    woNumber: workOrders.woNumber,
    location: workOrders.location,
    fromStreet: workOrders.fromStreet,
    toStreet: workOrders.toStreet,
    hasScan: sql<boolean>`${workOrders.scanFileKey} IS NOT NULL`.as('hasScan'),
    photoCount: sql<number>`(SELECT count(*) FROM photos p WHERE p.wo_id = ${workOrders.id})`.as('photoCount'),
    docCount: sql<number>`(SELECT count(*) FROM documents d WHERE d.org_id = ${orgId} AND d.wo_ids @> ARRAY[${workOrders.woNumber}]::text[] AND d.storage_key IS NOT NULL)`.as('docCount'),
  })
    .from(workOrders)
    .where(and(
      eq(workOrders.orgId, orgId),
      eq(workOrders.contractNum, contractNum),
      eq(workOrders.regionCode, regionCode),
      sql`(${workOrders.scanFileKey} IS NOT NULL
        OR EXISTS (SELECT 1 FROM photos p WHERE p.wo_id = ${workOrders.id})
        OR EXISTS (SELECT 1 FROM documents d WHERE d.org_id = ${orgId} AND d.wo_ids @> ARRAY[${workOrders.woNumber}]::text[] AND d.storage_key IS NOT NULL)
      )`,
    ))
    .orderBy(desc(workOrders.createdAt));

  return rows.map(r => {
    const fileCount = (r.hasScan ? 1 : 0) + Number(r.photoCount) + Number(r.docCount);
    const location = r.location || '';
    const streets = [r.fromStreet, r.toStreet].filter(Boolean).join(' → ');
    return {
      id: r.id,
      woNumber: r.woNumber,
      label: r.woNumber,
      sublabel: streets ? `${location} (${streets})` : location,
      count: fileCount,
    };
  });
}
