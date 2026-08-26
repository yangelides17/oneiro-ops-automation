export const JOB_TYPES = {
  FILL_PRODUCTION_LOG: 'fill_production_log',
  FILL_CERTIFIED_PAYROLL: 'fill_certified_payroll',
  FILL_SIGNIN: 'fill_signin',
  FILL_FIELD_REPORT: 'fill_field_report',
  FILL_MONTH_END: 'fill_month_end',
  SCAN_WORK_ORDER: 'scan_work_order',
  PROCESS_APPROVED_DOCS: 'process_approved_docs',
  SYNC_TO_DRIVE: 'sync_to_drive',
  SEND_EMAIL: 'send_email',
} as const;

export type JobType = typeof JOB_TYPES[keyof typeof JOB_TYPES];

export interface BaseJobPayload {
  orgId: string;
}

export interface FillJobPayload extends BaseJobPayload {
  documentId: string;         // documents table ID
  templateStorageKey: string; // R2 key for the template PDF
  fillData: Record<string, unknown>; // same JSON shape as current Drive JSONs
  overwriteStorageKey?: string; // if regenerating in place
}

export interface ScanJobPayload extends BaseJobPayload {
  storageKey: string;   // R2 key for the uploaded scan
  filename: string;
  mimeType: string;
}

export interface EmailJobPayload extends BaseJobPayload {
  to: string[];
  subject: string;
  html: string;
  attachmentKeys?: string[]; // R2 keys for attachments
}

export interface DriveSyncJobPayload extends BaseJobPayload {
  storageKey: string;
  filename: string;
  driveFolderPath: string;
}
