import { eq, and, asc, desc, sql, between, inArray } from 'drizzle-orm';
import type { Db } from '../client.js';
import { documents } from '../schema.js';

export async function getDocument(db: Db, orgId: string, id: string) {
  const [doc] = await db.select()
    .from(documents)
    .where(and(eq(documents.id, id), eq(documents.orgId, orgId)))
    .limit(1);
  return doc ?? null;
}

export async function getDocumentByKey(db: Db, orgId: string, docKey: string) {
  const [doc] = await db.select()
    .from(documents)
    .where(and(eq(documents.docKey, docKey), eq(documents.orgId, orgId)))
    .limit(1);
  return doc ?? null;
}

export async function listPendingDocuments(db: Db, orgId: string) {
  return db.select()
    .from(documents)
    .where(and(eq(documents.orgId, orgId), eq(documents.status, 'needs_review')))
    .orderBy(asc(documents.createdAt));
}

export async function listDocumentsByStatus(db: Db, orgId: string, status: string) {
  return db.select()
    .from(documents)
    .where(and(eq(documents.orgId, orgId), eq(documents.status, status)))
    .orderBy(desc(documents.createdAt));
}

export async function createDocument(db: Db, orgId: string, data: typeof documents.$inferInsert) {
  const [doc] = await db.insert(documents)
    .values({ ...data, orgId })
    .onConflictDoUpdate({
      target: [documents.orgId, documents.docKey],
      set: {
        storageKey: data.storageKey,
        filename: data.filename,
        status: data.status,
        updatedAt: new Date(),
      },
    })
    .returning();
  return doc;
}

export async function updateDocument(db: Db, orgId: string, id: string, data: Partial<typeof documents.$inferInsert>) {
  const [doc] = await db.update(documents)
    .set({ ...data, updatedAt: new Date() })
    .where(and(eq(documents.id, id), eq(documents.orgId, orgId)))
    .returning();
  return doc;
}

/** Set done/sent flags on a document. */
export async function setDocumentFlags(
  db: Db, orgId: string, id: string,
  flags: { done?: boolean; sent?: boolean },
) {
  const updates: Record<string, unknown> = { updatedAt: new Date() };
  if (flags.done !== undefined) {
    updates.done = flags.done;
    updates.doneAt = flags.done ? new Date() : null;
  }
  if (flags.sent !== undefined) {
    updates.sent = flags.sent;
    updates.sentAt = flags.sent ? new Date() : null;
  }
  const [doc] = await db.update(documents)
    .set(updates)
    .where(and(eq(documents.id, id), eq(documents.orgId, orgId)))
    .returning();
  return doc;
}

/** Get document status calendar data for a month. */
export async function getDocStatusCalendar(db: Db, orgId: string, monthStart: string, monthEnd: string) {
  return db.select()
    .from(documents)
    .where(and(
      eq(documents.orgId, orgId),
      between(documents.anchorDate, monthStart, monthEnd),
    ))
    .orderBy(asc(documents.anchorDate));
}

/** Get pending counts for nav badges. */
export async function getPendingCounts(db: Db, orgId: string) {
  const [result] = await db.select({
    approvals: sql<number>`count(*) filter (where ${documents.status} = 'needs_review')`,
    approved: sql<number>`count(*) filter (where ${documents.status} = 'approved')`,
    notDone: sql<number>`count(*) filter (where ${documents.done} = false and ${documents.status} != 'pending')`,
  })
    .from(documents)
    .where(eq(documents.orgId, orgId));

  // Sign-in pending count: distinct (date, contractNum, regionCode) groups in WDL
  const { workDayLog } = await import('../schema.js');
  const [wdlResult] = await db.select({
    count: sql<number>`count(distinct concat(${workDayLog.workDate}, '|', ${workDayLog.contractNum}, '|', ${workDayLog.regionCode}))`,
  })
    .from(workDayLog)
    .where(and(eq(workDayLog.orgId, orgId), eq(workDayLog.status, 'pending')));

  return {
    approvals_review: Number(result?.approvals ?? 0),
    approved_docs_pending: Number(result?.approved ?? 0),
    doc_status_pending: Number(result?.notDone ?? 0),
    signins_pending: Number(wdlResult?.count ?? 0),
  };
}

/** List documents for batch download, filtered by mode. */
export async function listDocumentsForBatch(
  db: Db, orgId: string,
  filters: {
    docTypes?: string[];
    contractorId?: string;
    dateStart?: string;
    dateEnd?: string;
    onlyUnsent?: boolean;
  },
) {
  const conditions = [eq(documents.orgId, orgId)];

  if (filters.docTypes?.length) {
    conditions.push(inArray(documents.docType, filters.docTypes));
  }
  if (filters.contractorId) {
    conditions.push(eq(documents.contractorId, filters.contractorId));
  }
  if (filters.dateStart && filters.dateEnd) {
    conditions.push(between(documents.anchorDate, filters.dateStart, filters.dateEnd));
  }
  if (filters.onlyUnsent) {
    conditions.push(eq(documents.done, true));
    conditions.push(eq(documents.sent, false));
  }

  return db.select()
    .from(documents)
    .where(and(...conditions))
    .orderBy(asc(documents.anchorDate))
    .limit(500);
}
