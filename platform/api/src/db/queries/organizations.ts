import { eq } from 'drizzle-orm';
import type { Db } from '../client.js';
import { organizations } from '../schema.js';

export async function getOrg(db: Db, orgId: string) {
  const [org] = await db.select()
    .from(organizations)
    .where(eq(organizations.id, orgId))
    .limit(1);
  return org ?? null;
}

export async function updateOrg(db: Db, orgId: string, data: Partial<typeof organizations.$inferInsert>) {
  const [updated] = await db.update(organizations)
    .set({ ...data, updatedAt: new Date() })
    .where(eq(organizations.id, orgId))
    .returning();
  return updated;
}
