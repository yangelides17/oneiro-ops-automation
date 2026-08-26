import { eq, and, asc, desc, sql } from 'drizzle-orm';
import type { Db } from '../client.js';
import { signinEntries, workDayLog } from '../schema.js';

// ─── Sign-In Entries ─────────────────────────────────────────

export async function listSigninByDate(db: Db, orgId: string, workDate: string) {
  return db.select()
    .from(signinEntries)
    .where(and(eq(signinEntries.orgId, orgId), eq(signinEntries.workDate, workDate)))
    .orderBy(asc(signinEntries.employeeName));
}

export async function listSigninByWo(db: Db, orgId: string, woId: string) {
  return db.select()
    .from(signinEntries)
    .where(and(eq(signinEntries.orgId, orgId), eq(signinEntries.woId, woId)))
    .orderBy(asc(signinEntries.workDate), asc(signinEntries.employeeName));
}

export async function createSigninEntry(db: Db, orgId: string, data: typeof signinEntries.$inferInsert) {
  const [entry] = await db.insert(signinEntries)
    .values({ ...data, orgId })
    .returning();
  return entry;
}

export async function createSigninEntriesBulk(db: Db, orgId: string, entries: typeof signinEntries.$inferInsert[]) {
  if (entries.length === 0) return [];
  return db.insert(signinEntries)
    .values(entries.map(e => ({ ...e, orgId })))
    .returning();
}

export async function updateSigninEntry(db: Db, orgId: string, id: string, data: Partial<typeof signinEntries.$inferInsert>) {
  const [entry] = await db.update(signinEntries)
    .set({ ...data, updatedAt: new Date() })
    .where(and(eq(signinEntries.id, id), eq(signinEntries.orgId, orgId)))
    .returning();
  return entry;
}

/**
 * Get hours summary for a date/contract/region/crew — used for OT lookback.
 * Returns total hours per employee on this date across all WOs.
 */
export async function getDayHoursSummary(
  db: Db, orgId: string,
  filters: { workDate: string; contractNum?: string; regionCode?: string; crewChief?: string },
) {
  let query = db.select({
    employeeName: signinEntries.employeeName,
    totalHours: sql<number>`sum(${signinEntries.hoursWorked})`.as('totalHours'),
    totalOt: sql<number>`sum(${signinEntries.otHours})`.as('totalOt'),
  })
    .from(signinEntries)
    .where(and(
      eq(signinEntries.orgId, orgId),
      eq(signinEntries.workDate, filters.workDate),
    ))
    .groupBy(signinEntries.employeeName);

  return query;
}

/**
 * Get sign-in rows associated with a specific document.
 * Matches by date + contract + region + optional crew chief.
 */
export async function getSigninRowsForDocument(
  db: Db, orgId: string,
  filters: { workDate: string; contractNum: string; regionCode: string; crewChief?: string },
) {
  const conditions = [
    eq(signinEntries.orgId, orgId),
    eq(signinEntries.workDate, filters.workDate),
    eq(signinEntries.contractNum, filters.contractNum),
    eq(signinEntries.regionCode, filters.regionCode),
  ];
  if (filters.crewChief) {
    conditions.push(eq(signinEntries.crewChief, filters.crewChief));
  }
  return db.select()
    .from(signinEntries)
    .where(and(...conditions))
    .orderBy(asc(signinEntries.employeeName));
}

// ─── Work Day Log ────────────────────────────────────────────

export async function listWorkDayLogQueue(db: Db, orgId: string) {
  return db.select()
    .from(workDayLog)
    .where(and(eq(workDayLog.orgId, orgId), eq(workDayLog.status, 'pending')))
    .orderBy(asc(workDayLog.workDate)); // FIFO: oldest first, surfaces overdue work
}

export async function createWorkDayLogEntry(db: Db, orgId: string, data: typeof workDayLog.$inferInsert) {
  const [entry] = await db.insert(workDayLog)
    .values({ ...data, orgId })
    .returning();
  return entry;
}

export async function updateWorkDayLogStatus(db: Db, orgId: string, id: string, status: string) {
  const [entry] = await db.update(workDayLog)
    .set({ status })
    .where(and(eq(workDayLog.id, id), eq(workDayLog.orgId, orgId)))
    .returning();
  return entry;
}
