import { eq, and } from 'drizzle-orm';
import type { Db } from '../client.js';
import { users } from '../schema.js';

export async function getUserById(db: Db, userId: string) {
  const [user] = await db.select()
    .from(users)
    .where(eq(users.id, userId))
    .limit(1);
  return user ?? null;
}

export async function getUserByEmail(db: Db, email: string) {
  const [user] = await db.select()
    .from(users)
    .where(eq(users.email, email.toLowerCase()))
    .limit(1);
  return user ?? null;
}

export async function listOrgUsers(db: Db, orgId: string) {
  return db.select({
    id: users.id,
    email: users.email,
    name: users.name,
    role: users.role,
    lastLoginAt: users.lastLoginAt,
    createdAt: users.createdAt,
  })
    .from(users)
    .where(eq(users.orgId, orgId))
    .orderBy(users.createdAt);
}

export async function updateUserRole(db: Db, orgId: string, userId: string, role: string) {
  const [updated] = await db.update(users)
    .set({ role })
    .where(and(eq(users.id, userId), eq(users.orgId, orgId)))
    .returning();
  return updated;
}
