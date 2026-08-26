import {
  pgTable, uuid, text, timestamp, integer, boolean, numeric,
  date, jsonb, bigserial, bigint, uniqueIndex, index,
  type AnyPgColumn,
} from 'drizzle-orm/pg-core';

// ─── Tenant & Auth ───────────────────────────────────────────

export const organizations = pgTable('organizations', {
  id: uuid('id').primaryKey().defaultRandom(),
  name: text('name').notNull(),
  address: text('address'),
  phone: text('phone'),
  email: text('email'),
  taxId: text('tax_id'),
  timezone: text('timezone').notNull().default('America/New_York'),
  opDayCutoffHour: integer('op_day_cutoff_hour').notNull().default(5),
  signatoryName: text('signatory_name'),
  signatoryTitle: text('signatory_title'),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
  updatedAt: timestamp('updated_at', { withTimezone: true }).notNull().defaultNow(),
});

export const users = pgTable('users', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  email: text('email').notNull().unique(),
  passwordHash: text('password_hash').notNull(),
  name: text('name').notNull(),
  role: text('role').notNull(),
  invitedBy: uuid('invited_by').references((): AnyPgColumn => users.id),
  invitedAt: timestamp('invited_at', { withTimezone: true }),
  lastLoginAt: timestamp('last_login_at', { withTimezone: true }),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
});

export const sessions = pgTable('sessions', {
  id: uuid('id').primaryKey().defaultRandom(),
  userId: uuid('user_id').notNull().references(() => users.id, { onDelete: 'cascade' }),
  tokenHash: text('token_hash').notNull(),
  expiresAt: timestamp('expires_at', { withTimezone: true }).notNull(),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
});

export const invitations = pgTable('invitations', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  email: text('email').notNull(),
  role: text('role').notNull(),
  invitedBy: uuid('invited_by').notNull().references(() => users.id),
  tokenHash: text('token_hash').notNull(),
  acceptedAt: timestamp('accepted_at', { withTimezone: true }),
  expiresAt: timestamp('expires_at', { withTimezone: true }).notNull(),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
});

// ─── Tenant Configuration ────────────────────────────────────

export const regions = pgTable('regions', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  code: text('code').notNull(),
  name: text('name').notNull(),
  sortOrder: integer('sort_order').notNull().default(0),
}, (t) => ({
  uniq: uniqueIndex('uq_regions_org_code').on(t.orgId, t.code),
}));

export const contractors = pgTable('contractors', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  name: text('name').notNull(),
  contactName: text('contact_name'),
  contactEmail: text('contact_email'),
  contactPhone: text('contact_phone'),
  address: text('address'),
  autoGeneratePl: boolean('auto_generate_pl').notNull().default(false),
  receivesPl: boolean('receives_pl').notNull().default(false),
  receivesCfr: boolean('receives_cfr').notNull().default(false),
  receivesInvoice: boolean('receives_invoice').notNull().default(false),
  receivesCp: boolean('receives_cp').notNull().default(false),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
}, (t) => ({
  uniq: uniqueIndex('uq_contractors_org_name').on(t.orgId, t.name),
}));

export const markingCategories = pgTable('marking_categories', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  name: text('name').notNull(),
  displayName: text('display_name'),
  unit: text('unit').notNull(),
  pricingGroup: text('pricing_group'),
  formSection: text('form_section').notNull().default('default'),
  requiresColor: boolean('requires_color').notNull().default(false),
  sortOrder: integer('sort_order').notNull().default(0),
  isActive: boolean('is_active').notNull().default(true),
}, (t) => ({
  uniq: uniqueIndex('uq_categories_org_name').on(t.orgId, t.name),
}));

export const pricingMultipliers = pgTable('pricing_multipliers', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  categoryName: text('category_name').notNull(),
  multiplierType: text('multiplier_type').notNull(),
  value: numeric('value', { precision: 10, scale: 4 }).notNull(),
}, (t) => ({
  uniq: uniqueIndex('uq_multipliers_org_cat_type').on(t.orgId, t.categoryName, t.multiplierType),
}));

export const contractPricing = pgTable('contract_pricing', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  contractorId: uuid('contractor_id').notNull().references(() => contractors.id),
  contractNum: text('contract_num').notNull(),
  regionCode: text('region_code'),
  effectiveDate: date('effective_date'),
  rateLine4: numeric('rate_line4', { precision: 10, scale: 4 }),
  rateLine12: numeric('rate_line12', { precision: 10, scale: 4 }),
  ratePreformed: numeric('rate_preformed', { precision: 10, scale: 4 }),
  rateExtruded: numeric('rate_extruded', { precision: 10, scale: 4 }),
  rateColorSurface: numeric('rate_color_surface', { precision: 10, scale: 4 }),
  notes: text('notes'),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
});

export const payClassifications = pgTable('pay_classifications', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  code: text('code').notNull(),
  name: text('name').notNull(),
  sortOrder: integer('sort_order').notNull().default(0),
}, (t) => ({
  uniq: uniqueIndex('uq_classifications_org_code').on(t.orgId, t.code),
}));

export const payRates = pgTable('pay_rates', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  classificationCode: text('classification_code').notNull(),
  effectiveDate: date('effective_date').notNull(),
  rateSt: numeric('rate_st', { precision: 10, scale: 2 }).notNull(),
  rateOt: numeric('rate_ot', { precision: 10, scale: 2 }).notNull(),
  suppSt: numeric('supp_st', { precision: 10, scale: 2 }).notNull().default('0'),
  suppOt: numeric('supp_ot', { precision: 10, scale: 2 }).notNull().default('0'),
  notes: text('notes'),
}, (t) => ({
  uniq: uniqueIndex('uq_pay_rates_org_class_date').on(t.orgId, t.classificationCode, t.effectiveDate),
}));

export const overtimeRules = pgTable('overtime_rules', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id).unique(),
  dailyThresholdHours: numeric('daily_threshold_hours', { precision: 4, scale: 2 }).default('8'),
  weeklyThresholdHours: numeric('weekly_threshold_hours', { precision: 4, scale: 2 }),
  weekendAllOt: boolean('weekend_all_ot').notNull().default(true),
  crossGroupLookback: boolean('cross_group_lookback').notNull().default(true),
});

export const billingRemaps = pgTable('billing_remaps', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  sourceContract: text('source_contract').notNull(),
  sourceRegion: text('source_region').notNull(),
  sourceContractor: text('source_contractor'),
  targetContract: text('target_contract').notNull(),
  targetRegion: text('target_region').notNull(),
  effectiveDate: date('effective_date').notNull(),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
});

export const employees = pgTable('employees', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  name: text('name').notNull(),
  address: text('address'),
  ssnLast4: text('ssn_last4'),
  raceEthnicity: text('race_ethnicity'),
  gender: text('gender'),
  isActive: boolean('is_active').notNull().default(true),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
}, (t) => ({
  uniq: uniqueIndex('uq_employees_org_name').on(t.orgId, t.name),
}));

export const contractLookup = pgTable('contract_lookup', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  contractNum: text('contract_num').notNull(),
  regionCode: text('region_code').notNull(),
  regionName: text('region_name'),
  contractId: text('contract_id'),
  projectName: text('project_name'),
}, (t) => ({
  uniq: uniqueIndex('uq_contract_lookup_org_num_region').on(t.orgId, t.contractNum, t.regionCode),
}));

// ─── Core Operational Tables ─────────────────────────────────

export const workOrders = pgTable('work_orders', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  woNumber: text('wo_number').notNull(),
  contractorId: uuid('contractor_id').notNull().references(() => contractors.id),
  contractNum: text('contract_num'),
  regionCode: text('region_code'),
  contractId: text('contract_id'),
  location: text('location'),
  fromStreet: text('from_street'),
  toStreet: text('to_street'),
  dueDate: date('due_date'),
  priority: text('priority'),
  workType: text('work_type'),
  woReceivedDate: date('wo_received_date'),
  waterBlastRequired: text('water_blast_required'),
  waterBlastConfirmed: text('water_blast_confirmed'),
  waterBlastSqft: numeric('water_blast_sqft'),
  status: text('status').notNull().default('received'),
  dispatchDate: date('dispatch_date'),
  workStartDate: date('work_start_date'),
  workEndDate: date('work_end_date'),
  issuesReported: text('issues_reported'),
  notes: text('notes'),
  dateEntered: date('date_entered'),
  school: text('school'),
  prepBy: text('prep_by'),
  generalRemarks: text('general_remarks'),
  latitude: numeric('latitude', { precision: 10, scale: 7 }),
  longitude: numeric('longitude', { precision: 10, scale: 7 }),
  geocodeWarning: text('geocode_warning'),
  geocodedAt: timestamp('geocoded_at', { withTimezone: true }),
  scanFileKey: text('scan_file_key'),
  scanCombinedId: text('scan_combined_id'),
  originalFilename: text('original_filename'),
  scanData: jsonb('scan_data'),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
  updatedAt: timestamp('updated_at', { withTimezone: true }).notNull().defaultNow(),
}, (t) => ({
  uniq: uniqueIndex('uq_wo_org_number').on(t.orgId, t.woNumber),
  statusIdx: index('idx_wo_org_status').on(t.orgId, t.status),
  contractorIdx: index('idx_wo_org_contractor').on(t.orgId, t.contractorId),
}));

export const markingItems = pgTable('marking_items', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  woId: uuid('wo_id').notNull().references(() => workOrders.id, { onDelete: 'cascade' }),
  workType: text('work_type'),
  woSection: text('wo_section'),
  category: text('category').notNull(),
  intersection: text('intersection'),
  direction: text('direction'),
  description: text('description'),
  quantity: numeric('quantity', { precision: 12, scale: 2 }),
  unit: text('unit'),
  colorMaterial: text('color_material'),
  dateCompleted: date('date_completed'),
  status: text('status').notNull().default('pending'),
  crewChief: text('crew_chief'),
  addedBy: text('added_by'),
  notes: text('notes'),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
  updatedAt: timestamp('updated_at', { withTimezone: true }).notNull().defaultNow(),
}, (t) => ({
  woIdx: index('idx_mi_wo').on(t.woId),
  statusIdx: index('idx_mi_org_status').on(t.orgId, t.status),
}));

export const signinEntries = pgTable('signin_entries', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  workDate: date('work_date').notNull(),
  woId: uuid('wo_id').notNull().references(() => workOrders.id),
  contractorId: uuid('contractor_id').notNull().references(() => contractors.id),
  contractNum: text('contract_num'),
  regionCode: text('region_code'),
  location: text('location'),
  employeeName: text('employee_name').notNull(),
  classification: text('classification').notNull(),
  timeIn: text('time_in'),
  timeOut: text('time_out'),
  hoursWorked: numeric('hours_worked', { precision: 5, scale: 2 }),
  otHours: numeric('ot_hours', { precision: 5, scale: 2 }),
  crewChief: text('crew_chief'),
  adminReviewed: boolean('admin_reviewed').notNull().default(false),
  reviewNotes: text('review_notes'),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
  updatedAt: timestamp('updated_at', { withTimezone: true }).notNull().defaultNow(),
}, (t) => ({
  dateIdx: index('idx_signin_org_date').on(t.orgId, t.workDate),
  woIdx: index('idx_signin_wo').on(t.woId),
}));

export const workDayLog = pgTable('work_day_log', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  workDate: date('work_date').notNull(),
  woId: uuid('wo_id').notNull().references(() => workOrders.id),
  contractorId: uuid('contractor_id').notNull().references(() => contractors.id),
  contractNum: text('contract_num'),
  regionCode: text('region_code'),
  location: text('location'),
  crewChief: text('crew_chief'),
  frSubmittedAt: timestamp('fr_submitted_at', { withTimezone: true }),
  status: text('status').notNull().default('pending'),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
}, (t) => ({
  dateIdx: index('idx_wdl_org_date').on(t.orgId, t.workDate),
}));

export const documents = pgTable('documents', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  docType: text('doc_type').notNull(),
  docKey: text('doc_key').notNull(),
  anchorDate: date('anchor_date'),
  contractorId: uuid('contractor_id').references(() => contractors.id),
  contractNum: text('contract_num'),
  regionCode: text('region_code'),
  woIds: text('wo_ids').array(),
  crewChief: text('crew_chief'),
  status: text('status').notNull().default('pending'),
  done: boolean('done').notNull().default(false),
  sent: boolean('sent').notNull().default(false),
  storageKey: text('storage_key'),
  filename: text('filename'),
  doneAt: timestamp('done_at', { withTimezone: true }),
  sentAt: timestamp('sent_at', { withTimezone: true }),
  notes: text('notes'),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
  updatedAt: timestamp('updated_at', { withTimezone: true }).notNull().defaultNow(),
}, (t) => ({
  uniq: uniqueIndex('uq_docs_org_key').on(t.orgId, t.docKey),
  statusIdx: index('idx_docs_org_status').on(t.orgId, t.status),
  typeDateIdx: index('idx_docs_org_type_date').on(t.orgId, t.docType, t.anchorDate),
}));

export const payrollEntries = pgTable('payroll_entries', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  weekStart: date('week_start').notNull(),
  weekEnd: date('week_end').notNull(),
  contractNum: text('contract_num').notNull(),
  regionCode: text('region_code'),
  contractId: text('contract_id'),
  projectName: text('project_name'),
  employeeName: text('employee_name').notNull(),
  classification: text('classification').notNull(),
  hoursByDay: jsonb('hours_by_day').notNull().default({ sun: 0, mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0 }),
  otByDay: jsonb('ot_by_day').notNull().default({ sun: 0, mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0 }),
  totalSt: numeric('total_st', { precision: 6, scale: 2 }).notNull().default('0'),
  totalOt: numeric('total_ot', { precision: 6, scale: 2 }).notNull().default('0'),
  rateSt: numeric('rate_st', { precision: 10, scale: 2 }),
  rateOt: numeric('rate_ot', { precision: 10, scale: 2 }),
  grossPay: numeric('gross_pay', { precision: 10, scale: 2 }),
  allWorkGross: numeric('all_work_gross', { precision: 10, scale: 2 }),
  deductions: numeric('deductions', { precision: 10, scale: 2 }),
  netPay: numeric('net_pay', { precision: 10, scale: 2 }),
  suppSt: numeric('supp_st', { precision: 10, scale: 2 }),
  suppOt: numeric('supp_ot', { precision: 10, scale: 2 }),
  matchStatus: text('match_status'),
  sentStatus: text('sent_status').default('No'),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
}, (t) => ({
  weekIdx: index('idx_payroll_org_week').on(t.orgId, t.weekStart),
}));

export const invoices = pgTable('invoices', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  invoiceNumber: text('invoice_number').notNull(),
  invoiceDate: date('invoice_date').notNull(),
  dueDate: date('due_date'),
  contractorId: uuid('contractor_id').references(() => contractors.id),
  contractNum: text('contract_num'),
  regionCode: text('region_code'),
  woId: uuid('wo_id').references(() => workOrders.id),
  description: text('description'),
  sqft: numeric('sqft', { precision: 12, scale: 2 }),
  rate: numeric('rate', { precision: 10, scale: 4 }),
  amount: numeric('amount', { precision: 12, scale: 2 }),
  status: text('status').notNull().default('draft'),
  qbInvoiceId: text('qb_invoice_id'),
  qbDocNumber: text('qb_doc_number'),
  paymentReceived: boolean('payment_received').notNull().default(false),
  paymentDate: date('payment_date'),
  notes: text('notes'),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
}, (t) => ({
  uniq: uniqueIndex('uq_invoices_org_number').on(t.orgId, t.invoiceNumber),
}));

export const photos = pgTable('photos', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  woId: uuid('wo_id').notNull().references(() => workOrders.id, { onDelete: 'cascade' }),
  storageKey: text('storage_key').notNull(),
  filename: text('filename').notNull(),
  mimeType: text('mime_type').notNull(),
  sizeBytes: bigint('size_bytes', { mode: 'number' }),
  thumbnailKey: text('thumbnail_key'),
  latitude: numeric('latitude', { precision: 10, scale: 7 }),
  longitude: numeric('longitude', { precision: 10, scale: 7 }),
  address: text('address'),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
}, (t) => ({
  woIdx: index('idx_photos_wo').on(t.woId),
}));

export const signatures = pgTable('signatures', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  woId: uuid('wo_id').references(() => workOrders.id),
  documentId: uuid('document_id').references(() => documents.id),
  storageKey: text('storage_key').notNull(),
  filename: text('filename').notNull(),
  signerName: text('signer_name'),
  signerTitle: text('signer_title'),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
});

export const auditLog = pgTable('audit_log', {
  id: bigserial('id', { mode: 'number' }).primaryKey(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  userId: uuid('user_id').references(() => users.id),
  source: text('source'),
  action: text('action').notNull(),
  subject: text('subject'),
  details: jsonb('details'),
  status: text('status'),
  actionRequired: text('action_required'),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
}, (t) => ({
  dateIdx: index('idx_audit_org_date').on(t.orgId, t.createdAt),
}));

// ─── Integration & Job Tables ────────────────────────────────

export const integrations = pgTable('integrations', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  type: text('type').notNull(),
  enabled: boolean('enabled').notNull().default(false),
  config: jsonb('config').notNull().default({}),
  credentials: text('credentials'),
  status: text('status').notNull().default('disconnected'),
  lastSyncAt: timestamp('last_sync_at', { withTimezone: true }),
  errorMessage: text('error_message'),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
  updatedAt: timestamp('updated_at', { withTimezone: true }).notNull().defaultNow(),
}, (t) => ({
  uniq: uniqueIndex('uq_integrations_org_type').on(t.orgId, t.type),
}));

export const jobs = pgTable('jobs', {
  id: uuid('id').primaryKey().defaultRandom(),
  orgId: uuid('org_id').notNull().references(() => organizations.id),
  type: text('type').notNull(),
  status: text('status').notNull().default('pending'),
  payload: jsonb('payload').notNull(),
  result: jsonb('result'),
  error: text('error'),
  attempts: integer('attempts').notNull().default(0),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
  completedAt: timestamp('completed_at', { withTimezone: true }),
}, (t) => ({
  orgStatusIdx: index('idx_jobs_org_status').on(t.orgId, t.status),
  typeStatusIdx: index('idx_jobs_type_status').on(t.type, t.status),
}));
