import { eq, and, asc, inArray } from 'drizzle-orm';
import type { Db } from '../client.js';
import { markingItems } from '../schema.js';

export async function listMarkingItems(db: Db, orgId: string, woId: string) {
  return db.select()
    .from(markingItems)
    .where(and(eq(markingItems.orgId, orgId), eq(markingItems.woId, woId)))
    .orderBy(asc(markingItems.createdAt));
}

export async function getMarkingItem(db: Db, orgId: string, id: string) {
  const [item] = await db.select()
    .from(markingItems)
    .where(and(eq(markingItems.id, id), eq(markingItems.orgId, orgId)))
    .limit(1);
  return item ?? null;
}

export async function createMarkingItem(db: Db, orgId: string, woId: string, data: { category: string } & Partial<typeof markingItems.$inferInsert>) {
  const [item] = await db.insert(markingItems)
    .values({ ...data, orgId, woId, category: data.category })
    .returning();
  return item;
}

export async function updateMarkingItem(db: Db, orgId: string, id: string, data: Partial<typeof markingItems.$inferInsert>) {
  const [item] = await db.update(markingItems)
    .set({ ...data, updatedAt: new Date() })
    .where(and(eq(markingItems.id, id), eq(markingItems.orgId, orgId)))
    .returning();
  return item;
}

export async function deleteMarkingItems(db: Db, orgId: string, ids: string[]) {
  if (ids.length === 0) return 0;
  const deleted = await db.delete(markingItems)
    .where(and(eq(markingItems.orgId, orgId), inArray(markingItems.id, ids)))
    .returning({ id: markingItems.id });
  return deleted.length;
}

/** Bulk create marking items (used when seeding from WO scan data). */
export async function createMarkingItemsBulk(db: Db, orgId: string, woId: string, items: ({ category: string } & Partial<typeof markingItems.$inferInsert>)[]) {
  if (items.length === 0) return [];
  return db.insert(markingItems)
    .values(items.map(item => ({ ...item, orgId, woId, category: item.category })))
    .returning();
}
