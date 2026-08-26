/**
 * Job Producers — enqueue fill, scan, email, and sync jobs.
 *
 * Each function inserts a row into the `jobs` table with the appropriate
 * type and payload. The Python worker polls for pending jobs and processes them.
 */

import { enqueueJob } from './queue.js';
import {
  JOB_TYPES,
  type FillJobPayload,
  type ScanJobPayload,
  type EmailJobPayload,
  type DriveSyncJobPayload,
} from './types.js';

export async function enqueueFillJob(type: string, payload: FillJobPayload) {
  try {
    const job = await enqueueJob(payload.orgId, type, payload as Record<string, unknown>);
    console.log(`[Queue] Fill job enqueued: ${type} (${job.id})`);
    return job;
  } catch (err: any) {
    console.error(`[Queue] Failed to enqueue fill job ${type}:`, err.message);
    return null;
  }
}

export async function enqueueScanJob(payload: ScanJobPayload) {
  try {
    return await enqueueJob(payload.orgId, JOB_TYPES.SCAN_WORK_ORDER, payload as Record<string, unknown>);
  } catch (err: any) {
    console.error('[Queue] Failed to enqueue scan job:', err.message);
    return null;
  }
}

export async function enqueueEmailJob(payload: EmailJobPayload) {
  try {
    return await enqueueJob(payload.orgId, JOB_TYPES.SEND_EMAIL, payload as Record<string, unknown>);
  } catch (err: any) {
    console.error('[Queue] Failed to enqueue email job:', err.message);
    return null;
  }
}

export async function enqueueDriveSyncJob(payload: DriveSyncJobPayload) {
  try {
    return await enqueueJob(payload.orgId, JOB_TYPES.SYNC_TO_DRIVE, payload as Record<string, unknown>);
  } catch (err: any) {
    console.error('[Queue] Failed to enqueue Drive sync job:', err.message);
    return null;
  }
}
