import { eq, and, desc } from 'drizzle-orm';
import type { Db } from '../client.js';
import { photos } from '../schema.js';

export async function listPhotos(db: Db, orgId: string, woId: string) {
  return db.select()
    .from(photos)
    .where(and(eq(photos.orgId, orgId), eq(photos.woId, woId)))
    .orderBy(desc(photos.createdAt))
    .limit(50);
}

export async function getPhoto(db: Db, orgId: string, id: string) {
  const [photo] = await db.select()
    .from(photos)
    .where(and(eq(photos.id, id), eq(photos.orgId, orgId)))
    .limit(1);
  return photo ?? null;
}

export async function createPhoto(db: Db, orgId: string, data: typeof photos.$inferInsert) {
  const [photo] = await db.insert(photos)
    .values({ ...data, orgId })
    .returning();
  return photo;
}

export async function deletePhoto(db: Db, orgId: string, id: string) {
  const [photo] = await db.delete(photos)
    .where(and(eq(photos.id, id), eq(photos.orgId, orgId)))
    .returning();
  return photo ?? null;
}
