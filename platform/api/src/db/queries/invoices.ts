import { eq, and, asc, desc, sql } from 'drizzle-orm';
import type { Db } from '../client.js';
import { invoices } from '../schema.js';

export async function listInvoices(db: Db, orgId: string) {
  return db.select()
    .from(invoices)
    .where(eq(invoices.orgId, orgId))
    .orderBy(desc(invoices.createdAt));
}

export async function getInvoice(db: Db, orgId: string, id: string) {
  const [inv] = await db.select()
    .from(invoices)
    .where(and(eq(invoices.id, id), eq(invoices.orgId, orgId)))
    .limit(1);
  return inv ?? null;
}

export async function getInvoiceByWo(db: Db, orgId: string, woId: string) {
  const [inv] = await db.select()
    .from(invoices)
    .where(and(eq(invoices.orgId, orgId), eq(invoices.woId, woId)))
    .limit(1);
  return inv ?? null;
}

export async function createInvoice(db: Db, orgId: string, data: typeof invoices.$inferInsert) {
  const [inv] = await db.insert(invoices)
    .values({ ...data, orgId })
    .returning();
  return inv;
}

export async function updateInvoice(db: Db, orgId: string, id: string, data: Partial<typeof invoices.$inferInsert>) {
  const [inv] = await db.update(invoices)
    .set(data)
    .where(and(eq(invoices.id, id), eq(invoices.orgId, orgId)))
    .returning();
  return inv;
}

/** Get the next invoice number for an org. */
export async function getNextInvoiceNumber(db: Db, orgId: string): Promise<string> {
  const [result] = await db.select({
    maxNum: sql<string>`max(${invoices.invoiceNumber})`,
  })
    .from(invoices)
    .where(eq(invoices.orgId, orgId));

  const current = parseInt(result?.maxNum || '1100', 10);
  return String(current + 1);
}
