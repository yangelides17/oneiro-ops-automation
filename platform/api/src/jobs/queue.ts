/**
 * Job Queue — PostgreSQL-backed
 *
 * Replaces the Redis/BullMQ queue with direct INSERT into the `jobs` table.
 * The Python worker polls this table for pending jobs.
 *
 * This is the standard database-backed job queue pattern (used by
 * Sidekiq, Graphile Worker, Celery with DB backend). Scales to
 * thousands of jobs/day on PostgreSQL. Migrate to Redis/BullMQ later
 * if real-time throughput demands it.
 */

import { db } from '../db/client.js';
import { jobs } from '../db/schema.js';

/**
 * Enqueue a job by inserting into the jobs table.
 * Returns the created job record.
 */
export async function enqueueJob(
  orgId: string,
  type: string,
  payload: Record<string, unknown>,
): Promise<{ id: string; type: string; status: string }> {
  const [job] = await db.insert(jobs)
    .values({
      orgId,
      type,
      payload,
      status: 'pending',
    })
    .returning({
      id: jobs.id,
      type: jobs.type,
      status: jobs.status,
    });
  return job;
}
