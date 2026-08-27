import { eq, and, asc, desc, sql, inArray } from 'drizzle-orm';
import type { Db } from '../client.js';
import { workOrders, markingItems, contractors, documents, invoices } from '../schema.js';

export async function listWorkOrders(db: Db, orgId: string) {
  return db.select()
    .from(workOrders)
    .where(eq(workOrders.orgId, orgId))
    .orderBy(desc(workOrders.createdAt));
}

export async function getWorkOrder(db: Db, orgId: string, id: string) {
  const [wo] = await db.select()
    .from(workOrders)
    .where(and(eq(workOrders.id, id), eq(workOrders.orgId, orgId)))
    .limit(1);
  return wo ?? null;
}

export async function getWorkOrderByNumber(db: Db, orgId: string, woNumber: string) {
  const [wo] = await db.select()
    .from(workOrders)
    .where(and(eq(workOrders.woNumber, woNumber), eq(workOrders.orgId, orgId)))
    .limit(1);
  return wo ?? null;
}

export async function createWorkOrder(db: Db, orgId: string, data: typeof workOrders.$inferInsert) {
  const [wo] = await db.insert(workOrders)
    .values({ ...data, orgId })
    .returning();
  return wo;
}

export async function updateWorkOrder(
  db: Db, orgId: string, id: string,
  data: Partial<typeof workOrders.$inferInsert>,
) {
  const [wo] = await db.update(workOrders)
    .set({ ...data, updatedAt: new Date() })
    .where(and(eq(workOrders.id, id), eq(workOrders.orgId, orgId)))
    .returning();
  return wo;
}

export async function deleteWorkOrder(db: Db, orgId: string, id: string) {
  await db.delete(workOrders)
    .where(and(eq(workOrders.id, id), eq(workOrders.orgId, orgId)));
}

/** WOs with coordinates for map view — includes contractor name + assignee. */
export async function listWorkOrdersForMap(db: Db, orgId: string) {
  const assigneeAlias = sql`(SELECT name FROM users WHERE id = ${workOrders.assignedTo})`.as('assignedToName');
  return db.select({
    id: workOrders.id,
    woNumber: workOrders.woNumber,
    contractorName: contractors.name,
    contractorId: workOrders.contractorId,
    contractNum: workOrders.contractNum,
    regionCode: workOrders.regionCode,
    location: workOrders.location,
    fromStreet: workOrders.fromStreet,
    toStreet: workOrders.toStreet,
    status: workOrders.status,
    workType: workOrders.workType,
    priority: workOrders.priority,
    dueDate: workOrders.dueDate,
    latitude: workOrders.latitude,
    longitude: workOrders.longitude,
    geocodeWarning: workOrders.geocodeWarning,
    scanFileKey: workOrders.scanFileKey,
    assignedTo: workOrders.assignedTo,
    assignedToName: assigneeAlias,
  })
    .from(workOrders)
    .leftJoin(contractors, eq(workOrders.contractorId, contractors.id))
    .where(eq(workOrders.orgId, orgId));
}

/** WOs assigned to a specific user — for "My Work" view. */
export async function listMyWork(db: Db, orgId: string, userId: string) {
  return db.select({
    id: workOrders.id,
    woNumber: workOrders.woNumber,
    contractorName: contractors.name,
    contractNum: workOrders.contractNum,
    regionCode: workOrders.regionCode,
    location: workOrders.location,
    fromStreet: workOrders.fromStreet,
    toStreet: workOrders.toStreet,
    status: workOrders.status,
    workType: workOrders.workType,
    priority: workOrders.priority,
    dueDate: workOrders.dueDate,
    latitude: workOrders.latitude,
    longitude: workOrders.longitude,
    scanFileKey: workOrders.scanFileKey,
  })
    .from(workOrders)
    .leftJoin(contractors, eq(workOrders.contractorId, contractors.id))
    .where(and(eq(workOrders.orgId, orgId), eq(workOrders.assignedTo, userId)))
    .orderBy(
      sql`CASE ${workOrders.status}
        WHEN 'in_progress' THEN 0
        WHEN 'dispatched' THEN 1
        WHEN 'received' THEN 2
        WHEN 'completed' THEN 3
        ELSE 4 END`,
    );
}

/**
 * Dashboard data: WOs with contractor names + marking item rollups.
 * Returns clean camelCase shapes.
 */
export async function getDashboardData(db: Db, orgId: string) {
  const rows = await db.select({
    id: workOrders.id,
    woNumber: workOrders.woNumber,
    contractorId: workOrders.contractorId,
    contractorName: contractors.name,
    contractNum: workOrders.contractNum,
    regionCode: workOrders.regionCode,
    contractId: workOrders.contractId,
    location: workOrders.location,
    fromStreet: workOrders.fromStreet,
    toStreet: workOrders.toStreet,
    dueDate: workOrders.dueDate,
    priority: workOrders.priority,
    workType: workOrders.workType,
    woReceivedDate: workOrders.woReceivedDate,
    waterBlastRequired: workOrders.waterBlastRequired,
    waterBlastConfirmed: workOrders.waterBlastConfirmed,
    status: workOrders.status,
    dispatchDate: workOrders.dispatchDate,
    workStartDate: workOrders.workStartDate,
    workEndDate: workOrders.workEndDate,
    issuesReported: workOrders.issuesReported,
    notes: workOrders.notes,
    latitude: workOrders.latitude,
    longitude: workOrders.longitude,
    geocodeWarning: workOrders.geocodeWarning,
    createdAt: workOrders.createdAt,
  })
    .from(workOrders)
    .leftJoin(contractors, eq(workOrders.contractorId, contractors.id))
    .where(eq(workOrders.orgId, orgId))
    .orderBy(
      sql`CASE ${workOrders.status}
        WHEN 'in_progress' THEN 0
        WHEN 'dispatched'  THEN 1
        WHEN 'received'    THEN 2
        WHEN 'completed'   THEN 3
        WHEN 'returned'    THEN 4
        ELSE 5 END`,
      desc(workOrders.createdAt),
    );

  const emptyStats = { total: 0, received: 0, dispatched: 0, inProgress: 0, completed: 0, complete: 0 };
  if (rows.length === 0) {
    return { wos: [], stats: emptyStats, byContractor: {}, attention: [] };
  }

  const woIds = rows.map(w => w.id);

  // ── Marking item rollups ────────────────────────────────────
  const itemRollups = await db.select({
    woId: markingItems.woId,
    total: sql<number>`count(*)`.as('total'),
    completed: sql<number>`count(*) filter (where ${markingItems.status} = 'completed')`.as('completed'),
    totalQty: sql<number>`coalesce(sum(case when ${markingItems.status} = 'completed' then cast(${markingItems.quantity} as numeric) else 0 end), 0)`.as('totalQty'),
    primaryUnit: sql<string>`mode() within group (order by ${markingItems.unit})`.as('primaryUnit'),
  })
    .from(markingItems)
    .where(and(eq(markingItems.orgId, orgId), inArray(markingItems.woId, woIds)))
    .groupBy(markingItems.woId);

  const rollupMap = new Map(itemRollups.map(i => [i.woId, {
    total: Number(i.total),
    completed: Number(i.completed),
    qty: Number(i.totalQty),
    unit: i.primaryUnit || 'LF',
  }]));

  // ── Document lifecycle flags per WO ─────────────────────────
  // Build docs{} object per WO from documents table.
  // Old app had: cfr, production_log, signin, certified_payroll, invoice
  const docRows = await db.select({
    docType: documents.docType,
    woIds: documents.woIds,
    done: documents.done,
    sent: documents.sent,
  })
    .from(documents)
    .where(eq(documents.orgId, orgId));

  // Map WO number → { docType → { done, sent } }
  const docsMap = new Map<string, Record<string, { done: boolean; sent: boolean }>>();
  for (const doc of docRows) {
    for (const woNum of (doc.woIds || [])) {
      if (!docsMap.has(woNum)) docsMap.set(woNum, {});
      const entry = docsMap.get(woNum)!;
      const existing = entry[doc.docType];
      // If any doc of this type for this WO is done, the flag is true
      entry[doc.docType] = {
        done: existing?.done || doc.done,
        sent: existing?.sent || doc.sent,
      };
    }
  }

  // ── Invoice data per WO ─────────────────────────────────────
  const invoiceRows = await db.select({
    woId: invoices.woId,
    invoiceNumber: invoices.invoiceNumber,
    invoiceDate: invoices.invoiceDate,
    amount: invoices.amount,
  })
    .from(invoices)
    .where(eq(invoices.orgId, orgId));

  const invoiceMap = new Map(invoiceRows.map(inv => [inv.woId, inv]));

  // ── Assemble WOs ────────────────────────────────────────────
  const byContractor: Record<string, number> = {};
  const wos = rows.map(r => {
    const name = r.contractorName || 'Unknown';
    byContractor[name] = (byContractor[name] || 0) + 1;

    const rollup = rollupMap.get(r.id);
    const docFlags = docsMap.get(r.woNumber!) || {};
    const inv = invoiceMap.get(r.id);
    const isMma = (r.workType || '').toUpperCase() === 'MMA';

    return {
      ...r,
      contractorName: name,
      markingsTotal: rollup?.total ?? 0,
      markingsCompleted: rollup?.completed ?? 0,
      quantity: rollup ? String(Math.round(rollup.qty)) : '',
      quantityUnit: rollup?.unit || (isMma ? 'SF' : 'LF'),
      // Per-doc lifecycle flags (matches old app shape)
      docs: {
        cfr:               docFlags['field_report']      || { done: false, sent: false },
        production_log:    docFlags['production_log']    || { done: false, sent: false },
        signin:            docFlags['signin']            || { done: false, sent: false },
        certified_payroll: docFlags['certified_payroll'] || { done: false, sent: false },
        invoice:           inv ? { done: true, sent: false } : { done: false, sent: false },
      },
      // Invoice fields
      invoiceDocNumber: inv?.invoiceNumber || null,
      invoiceDate: inv?.invoiceDate || null,
      invoiceAmount: inv?.amount ? Number(inv.amount) : null,
    };
  });

  // ── Attention WOs ───────────────────────────────────────────
  // Matches Code.js: not completed AND (has issues OR in-progress without photos)
  const attention = wos
    .filter(w => w.status !== 'completed' && (
      (w.issuesReported && String(w.issuesReported).trim()) ||
      (w.status === 'in_progress' && !w.docs.cfr.done)
    ))
    .map(w => w.woNumber);

  // ── Stats (both camelCase and old-app keys) ─────────────────
  const completedCount = rows.filter(w => w.status === 'completed').length;
  const stats = {
    total: rows.length,
    received: rows.filter(w => w.status === 'received').length,
    dispatched: rows.filter(w => w.status === 'dispatched').length,
    inProgress: rows.filter(w => w.status === 'in_progress').length,
    completed: completedCount,
    complete: completedCount,  // old-app alias
  };

  return { wos, stats, byContractor, attention };
}

/** List WOs with contractor names — for dropdown/list use. */
export async function listWorkOrdersWithContractor(db: Db, orgId: string) {
  return db.select({
    id: workOrders.id,
    woNumber: workOrders.woNumber,
    contractorName: contractors.name,
    contractorId: workOrders.contractorId,
    contractNum: workOrders.contractNum,
    regionCode: workOrders.regionCode,
    location: workOrders.location,
    fromStreet: workOrders.fromStreet,
    toStreet: workOrders.toStreet,
    dueDate: workOrders.dueDate,
    priority: workOrders.priority,
    workType: workOrders.workType,
    status: workOrders.status,
    woReceivedDate: workOrders.woReceivedDate,
    waterBlastRequired: workOrders.waterBlastRequired,
    waterBlastConfirmed: workOrders.waterBlastConfirmed,
    dispatchDate: workOrders.dispatchDate,
    workStartDate: workOrders.workStartDate,
    workEndDate: workOrders.workEndDate,
    issuesReported: workOrders.issuesReported,
    notes: workOrders.notes,
    generalRemarks: workOrders.generalRemarks,
    school: workOrders.school,
    prepBy: workOrders.prepBy,
    dateEntered: workOrders.dateEntered,
    latitude: workOrders.latitude,
    longitude: workOrders.longitude,
    geocodeWarning: workOrders.geocodeWarning,
    scanFileKey: workOrders.scanFileKey,
    originalFilename: workOrders.originalFilename,
    createdAt: workOrders.createdAt,
  })
    .from(workOrders)
    .leftJoin(contractors, eq(workOrders.contractorId, contractors.id))
    .where(eq(workOrders.orgId, orgId))
    .orderBy(
      // Status priority: active WOs first, completed/returned last.
      // Matches Code.js handleGetActiveWOs_ sort order (lines 10959–10998).
      sql`CASE ${workOrders.status}
        WHEN 'in_progress' THEN 0
        WHEN 'dispatched'  THEN 1
        WHEN 'received'    THEN 2
        WHEN 'completed'   THEN 3
        WHEN 'returned'    THEN 4
        ELSE 5 END`,
      desc(workOrders.createdAt),
    );
}
