import { eq, desc } from 'drizzle-orm';
import type { Db } from '../client.js';
import { auditLog } from '../schema.js';

export async function createAuditEntry(
  db: Db, orgId: string,
  data: {
    userId?: string;
    source?: string;
    action: string;
    subject?: string;
    details?: Record<string, unknown>;
    status?: string;
    actionRequired?: string;
  },
) {
  const [entry] = await db.insert(auditLog)
    .values({ orgId, ...data })
    .returning();
  return entry;
}

export async function listRecentAudit(db: Db, orgId: string, limit = 100) {
  return db.select()
    .from(auditLog)
    .where(eq(auditLog.orgId, orgId))
    .orderBy(desc(auditLog.createdAt))
    .limit(limit);
}
