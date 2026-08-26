import { eq, and, asc } from 'drizzle-orm';
import type { Db } from '../client.js';
import { payrollEntries } from '../schema.js';

export async function listPayrollByWeek(db: Db, orgId: string, weekStart: string) {
  return db.select()
    .from(payrollEntries)
    .where(and(eq(payrollEntries.orgId, orgId), eq(payrollEntries.weekStart, weekStart)))
    .orderBy(asc(payrollEntries.employeeName));
}

export async function createPayrollEntry(db: Db, orgId: string, data: typeof payrollEntries.$inferInsert) {
  const [entry] = await db.insert(payrollEntries)
    .values({ ...data, orgId })
    .returning();
  return entry;
}

export async function createPayrollEntriesBulk(db: Db, orgId: string, entries: typeof payrollEntries.$inferInsert[]) {
  if (entries.length === 0) return [];
  return db.insert(payrollEntries)
    .values(entries.map(e => ({ ...e, orgId })))
    .returning();
}

export async function updatePayrollEntry(db: Db, orgId: string, id: string, data: Partial<typeof payrollEntries.$inferInsert>) {
  const [entry] = await db.update(payrollEntries)
    .set(data)
    .where(and(eq(payrollEntries.id, id), eq(payrollEntries.orgId, orgId)))
    .returning();
  return entry;
}
