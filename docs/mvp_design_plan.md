# MVP Design Plan — Multi-Tenant Rebuild

## Overview

Rebuild the Oneiro internal operations tool as a multi-tenant SaaS product. Same features, same workflows, same UI — new foundation that supports multiple companies.

### What Changes
- Google Sheets → PostgreSQL
- Apps Script API → TypeScript/Node backend
- Google Drive (as message bus) → Redis job queue + R2 object storage
- Shared door codes → Email/password auth with roles
- Hardcoded Oneiro config → Per-tenant configuration
- Tightly coupled integrations → Pluggable connector interfaces

### What Stays the Same
- All 9 React pages and 22 components (same UI, swapped API calls)
- All 10 client-side lib modules (pricing.js, signinShared.js, etc.)
- Python PDF fill pipeline (pypdf + PyMuPDF + Pillow + reportlab)
- Claude Vision WO scanning (same prompts, same model)
- Every user-facing workflow

---

## System Architecture

```
┌──────────────────────────────────────────────────────────────┐
│                        Clients (Browser)                      │
│                     React 18 SPA (Vite)                       │
└────────────────────────────┬─────────────────────────────────┘
                             │ HTTPS (JWT in HttpOnly cookie)
┌────────────────────────────▼─────────────────────────────────┐
│                     API Server (TypeScript/Node)              │
│                                                              │
│  Auth Middleware → Tenant Context → Role Check → Route       │
│                                                              │
│  ┌─────────────────────────────────────────────────────────┐ │
│  │              Business Logic Layer                       │ │
│  │  (pricing, OT calc, marking aggregation,                │ │
│  │   doc lifecycle, payroll, revenue — SINGLE SOURCE)      │ │
│  └─────────────────────┬───────────────────────────────────┘ │
│                        │                                     │
│  ┌─────────────────────┴───────────────────────────────────┐ │
│  │              Integration Gateway                        │ │
│  │  Calls connector interfaces, never external APIs direct │ │
│  └─────────────────────────────────────────────────────────┘ │
└────────────────────────┬─────────────────────────────────────┘
                         │
         ┌───────────────┼───────────────────┐
         │               │                   │
    PostgreSQL        Redis             Cloudflare R2
    (all data)     (job queue +         (PDFs, photos,
                    cache)              templates)
                         │
                   ┌─────▼──────────────────────────────┐
                   │        PDF Worker (Python)          │
                   │   Consumes jobs from Redis          │
                   │   Same fill_*.py code               │
                   │   Claude Vision for WO scanning     │
                   └────────────────────────────────────┘

         ┌──────────────────────────────────────┐
         │       Integration Connectors         │
         │  (optional, per-tenant)              │
         │                                      │
         │  Google Drive  (one-way push)        │
         │  QuickBooks    (deferred)            │
         │  Resend        (platform email)      │
         │  Google Maps   (platform geocoding)  │
         └──────────────────────────────────────┘
```

### Deployable Services (Railway)

| Service | Language | Role |
|---------|----------|------|
| **api** | TypeScript/Node | HTTP API + serves React SPA |
| **worker** | Python | Consumes doc fill + WO scan jobs |
| **postgres** | Managed | All application data |
| **redis** | Managed | Job queue (BullMQ) + ephemeral cache |

### External Services
- **Cloudflare R2** — document/photo storage (S3-compatible, zero egress fees)
- **Resend** — transactional email (platform account)
- **Google Maps API** — geocoding (platform key)
- **Anthropic Claude API** — WO scanning (platform key, opaque to tenants)

---

## Database Schema

### Tenant & Auth

```sql
CREATE TABLE organizations (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name                  TEXT NOT NULL,
  address               TEXT,
  phone                 TEXT,
  email                 TEXT,
  tax_id                TEXT,                             -- EIN (encrypted at app level)
  timezone              TEXT NOT NULL DEFAULT 'America/New_York',
  op_day_cutoff_hour    INT NOT NULL DEFAULT 5,
  signatory_name        TEXT,
  signatory_title       TEXT,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE users (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  email                 TEXT NOT NULL UNIQUE,
  password_hash         TEXT NOT NULL,
  name                  TEXT NOT NULL,
  role                  TEXT NOT NULL CHECK (role IN ('owner','admin','foreman','crew')),
  invited_by            UUID REFERENCES users(id),
  invited_at            TIMESTAMPTZ,
  last_login_at         TIMESTAMPTZ,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE sessions (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id               UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  token_hash            TEXT NOT NULL,
  expires_at            TIMESTAMPTZ NOT NULL,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE invitations (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  email                 TEXT NOT NULL,
  role                  TEXT NOT NULL CHECK (role IN ('admin','foreman','crew')),
  invited_by            UUID NOT NULL REFERENCES users(id),
  token_hash            TEXT NOT NULL,
  accepted_at           TIMESTAMPTZ,
  expires_at            TIMESTAMPTZ NOT NULL,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);
```

### Tenant Configuration

```sql
-- Regions (replaces hardcoded boroughs: MN, BK, BX, QU, SI)
CREATE TABLE regions (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  code                  TEXT NOT NULL,
  name                  TEXT NOT NULL,
  sort_order            INT NOT NULL DEFAULT 0,
  UNIQUE(org_id, code)
);

-- Prime contractors the org works under
CREATE TABLE contractors (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  name                  TEXT NOT NULL,
  contact_name          TEXT,
  contact_email         TEXT,
  contact_phone         TEXT,
  address               TEXT,
  auto_generate_pl      BOOLEAN NOT NULL DEFAULT false,
  receives_pl           BOOLEAN NOT NULL DEFAULT false,
  receives_cfr          BOOLEAN NOT NULL DEFAULT false,
  receives_invoice      BOOLEAN NOT NULL DEFAULT false,
  receives_cp           BOOLEAN NOT NULL DEFAULT false,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE(org_id, name)
);

-- 48 NYC DOT categories as default, add/remove per tenant
CREATE TABLE marking_categories (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  name                  TEXT NOT NULL,
  display_name          TEXT,
  unit                  TEXT NOT NULL CHECK (unit IN ('SF','LF','EA')),
  pricing_group         TEXT,
  form_section          TEXT NOT NULL DEFAULT 'default' CHECK (form_section IN ('grid','mma','default')),
  requires_color        BOOLEAN NOT NULL DEFAULT false,
  sort_order            INT NOT NULL DEFAULT 0,
  is_active             BOOLEAN NOT NULL DEFAULT true,
  UNIQUE(org_id, name)
);

-- Line width multipliers, extruded unit counts, preformed unit counts, line12 multipliers
CREATE TABLE pricing_multipliers (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  category_name         TEXT NOT NULL,
  multiplier_type       TEXT NOT NULL CHECK (multiplier_type IN (
                          'line_width','line12','extruded_unit','preformed_unit'
                        )),
  value                 NUMERIC(10,4) NOT NULL,
  UNIQUE(org_id, category_name, multiplier_type)
);

-- Rate tables: (contractor, contract#, region, effective date) → rates per pricing group
CREATE TABLE contract_pricing (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  contractor_id         UUID NOT NULL REFERENCES contractors(id),
  contract_num          TEXT NOT NULL,
  region_code           TEXT,
  effective_date        DATE,
  rate_line4            NUMERIC(10,4),
  rate_line12           NUMERIC(10,4),
  rate_preformed        NUMERIC(10,4),
  rate_extruded         NUMERIC(10,4),
  rate_color_surface    NUMERIC(10,4),
  notes                 TEXT,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Employee classifications (LP, SAT, OP, LAB, FGL, SUP)
CREATE TABLE pay_classifications (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  code                  TEXT NOT NULL,
  name                  TEXT NOT NULL,
  sort_order            INT NOT NULL DEFAULT 0,
  UNIQUE(org_id, code)
);

-- Pay rate schedule per classification with effective dates
CREATE TABLE pay_rates (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  classification_code   TEXT NOT NULL,
  effective_date        DATE NOT NULL,
  rate_st               NUMERIC(10,2) NOT NULL,
  rate_ot               NUMERIC(10,2) NOT NULL,
  supp_st               NUMERIC(10,2) NOT NULL DEFAULT 0,
  supp_ot               NUMERIC(10,2) NOT NULL DEFAULT 0,
  notes                 TEXT,
  UNIQUE(org_id, classification_code, effective_date)
);

-- Configurable OT rules per tenant
CREATE TABLE overtime_rules (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id) UNIQUE,
  daily_threshold_hours   NUMERIC(4,2) DEFAULT 8,
  weekly_threshold_hours  NUMERIC(4,2),
  weekend_all_ot          BOOLEAN NOT NULL DEFAULT true,
  cross_group_lookback    BOOLEAN NOT NULL DEFAULT true
);

-- Billing remap rules (contract reassignment for sub-primes)
CREATE TABLE billing_remaps (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  source_contract       TEXT NOT NULL,
  source_region         TEXT NOT NULL,
  source_contractor     TEXT,
  target_contract       TEXT NOT NULL,
  target_region         TEXT NOT NULL,
  effective_date        DATE NOT NULL,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE(org_id, source_contract, source_region, source_contractor, effective_date)
);

-- Employee registry
CREATE TABLE employees (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  name                  TEXT NOT NULL,
  address               TEXT,
  ssn_last4             TEXT,                             -- encrypted at app level
  race_ethnicity        TEXT,                             -- EEO categories (B/H/A/NA)
  gender                TEXT,                             -- for EU form (M/F)
  is_active             BOOLEAN NOT NULL DEFAULT true,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE(org_id, name)
);

-- Maps contract+region → Contract ID (Comptroller's Reg #) for certified payroll forms
CREATE TABLE contract_lookup (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  contract_num          TEXT NOT NULL,
  region_code           TEXT NOT NULL,
  region_name           TEXT,
  contract_id           TEXT,
  project_name          TEXT,
  UNIQUE(org_id, contract_num, region_code)
);
```

### Core Operational Tables

```sql
-- Work Orders (mirrors WO Tracker — 49 columns mapped to proper schema)
CREATE TABLE work_orders (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  wo_number             TEXT NOT NULL,
  contractor_id         UUID NOT NULL REFERENCES contractors(id),
  contract_num          TEXT,
  region_code           TEXT,
  contract_id           TEXT,                             -- looked up from contract_lookup
  location              TEXT,
  from_street           TEXT,
  to_street             TEXT,
  due_date              DATE,
  priority              TEXT,
  work_type             TEXT,                             -- "mma", "thermo"
  wo_received_date      DATE,
  water_blast_required  TEXT,
  water_blast_confirmed TEXT,
  water_blast_sqft      NUMERIC,
  status                TEXT NOT NULL DEFAULT 'received'
                          CHECK (status IN ('received','dispatched','in_progress','completed','returned')),
  dispatch_date         DATE,
  work_start_date       DATE,
  work_end_date         DATE,
  issues_reported       TEXT,
  notes                 TEXT,
  date_entered          DATE,                             -- from Vision parser (when WO was entered)
  school                TEXT,                             -- from Vision parser
  prep_by               TEXT,                             -- from Vision parser
  general_remarks       TEXT,
  latitude              NUMERIC(10,7),
  longitude             NUMERIC(10,7),
  geocode_warning       TEXT,
  geocoded_at           TIMESTAMPTZ,
  -- Scan tracking
  scan_file_key         TEXT,                             -- R2 storage key for scanned WO PDF
  scan_combined_id      TEXT,                             -- links splits from same multi-WO stack
  original_filename     TEXT,                             -- source filename from upload
  scan_data             JSONB,                            -- raw Vision output: {top_markings, intersection_grid, bike_lane_markings}
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE(org_id, wo_number)
);

CREATE INDEX idx_wo_org_status ON work_orders(org_id, status);
CREATE INDEX idx_wo_org_contractor ON work_orders(org_id, contractor_id);

-- Marking Items (mirrors 16-column Marking Items sheet)
CREATE TABLE marking_items (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  wo_id                 UUID NOT NULL REFERENCES work_orders(id) ON DELETE CASCADE,
  work_type             TEXT,
  wo_section            TEXT CHECK (wo_section IN ('top_table','intersection_grid','manual')),
  category              TEXT NOT NULL,
  intersection          TEXT,
  direction             TEXT,
  description           TEXT,
  quantity              NUMERIC(12,2),
  unit                  TEXT CHECK (unit IN ('SF','LF','EA')),
  color_material        TEXT,
  date_completed        DATE,
  status                TEXT NOT NULL DEFAULT 'pending'
                          CHECK (status IN ('pending','completed','skipped')),
  crew_chief            TEXT,
  added_by              TEXT CHECK (added_by IN ('scanner','manual')),
  notes                 TEXT,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_mi_wo ON marking_items(wo_id);
CREATE INDEX idx_mi_org_status ON marking_items(org_id, status);

-- Daily Sign-In Data (mirrors 15-column Daily Sign-In Data sheet)
CREATE TABLE signin_entries (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  work_date             DATE NOT NULL,
  wo_id                 UUID NOT NULL REFERENCES work_orders(id),
  contractor_id         UUID NOT NULL REFERENCES contractors(id),
  contract_num          TEXT,
  region_code           TEXT,
  location              TEXT,
  employee_name         TEXT NOT NULL,
  classification        TEXT NOT NULL,
  time_in               TEXT,
  time_out              TEXT,
  hours_worked          NUMERIC(5,2),
  ot_hours              NUMERIC(5,2),
  crew_chief            TEXT,
  admin_reviewed        BOOLEAN NOT NULL DEFAULT false,
  review_notes          TEXT,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_signin_org_date ON signin_entries(org_id, work_date);
CREATE INDEX idx_signin_wo ON signin_entries(wo_id);

-- Work Day Log (queue for sign-in filing — mirrors 9-column Work Day Log sheet)
CREATE TABLE work_day_log (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  work_date             DATE NOT NULL,
  wo_id                 UUID NOT NULL REFERENCES work_orders(id),
  contractor_id         UUID NOT NULL REFERENCES contractors(id),
  contract_num          TEXT,
  region_code           TEXT,
  location              TEXT,
  crew_chief            TEXT,
  fr_submitted_at       TIMESTAMPTZ,
  status                TEXT NOT NULL DEFAULT 'pending'
                          CHECK (status IN ('pending','submitted','skipped')),
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_wdl_org_date ON work_day_log(org_id, work_date);

-- Document Lifecycle (mirrors 13-column Doc Lifecycle Log)
CREATE TABLE documents (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  doc_type              TEXT NOT NULL CHECK (doc_type IN (
                          'production_log','signin','certified_payroll',
                          'field_report','employee_utilization','certificates'
                        )),
  doc_key               TEXT NOT NULL,                    -- synthetic dedup key
  anchor_date           DATE,                             -- work date, week start, or month
  contractor_id         UUID REFERENCES contractors(id),
  contract_num          TEXT,
  region_code           TEXT,
  wo_ids                TEXT[],
  crew_chief            TEXT,
  status                TEXT NOT NULL DEFAULT 'pending'
                          CHECK (status IN ('pending','needs_review','approved','archived')),
  done                  BOOLEAN NOT NULL DEFAULT false,
  sent                  BOOLEAN NOT NULL DEFAULT false,
  storage_key           TEXT,                             -- R2 key for the filled PDF
  filename              TEXT,
  done_at               TIMESTAMPTZ,
  sent_at               TIMESTAMPTZ,
  notes                 TEXT,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE(org_id, doc_key)
);

CREATE INDEX idx_docs_org_status ON documents(org_id, status);
CREATE INDEX idx_docs_org_type_date ON documents(org_id, doc_type, anchor_date);

-- Certified Payroll Tracker (mirrors 27-column Certified Payroll Tracker)
CREATE TABLE payroll_entries (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  week_start            DATE NOT NULL,
  week_end              DATE NOT NULL,
  contract_num          TEXT NOT NULL,
  region_code           TEXT,
  contract_id           TEXT,
  project_name          TEXT,
  employee_name         TEXT NOT NULL,
  classification        TEXT NOT NULL,
  hours_by_day          JSONB NOT NULL DEFAULT '{"sun":0,"mon":0,"tue":0,"wed":0,"thu":0,"fri":0,"sat":0}',
  ot_by_day             JSONB NOT NULL DEFAULT '{"sun":0,"mon":0,"tue":0,"wed":0,"thu":0,"fri":0,"sat":0}',
  total_st              NUMERIC(6,2) NOT NULL DEFAULT 0,
  total_ot              NUMERIC(6,2) NOT NULL DEFAULT 0,
  rate_st               NUMERIC(10,2),
  rate_ot               NUMERIC(10,2),
  gross_pay             NUMERIC(10,2),
  all_work_gross        NUMERIC(10,2),                    -- from paystub (all contracts, not just this)
  deductions            NUMERIC(10,2),
  net_pay               NUMERIC(10,2),
  supp_st               NUMERIC(10,2),
  supp_ot               NUMERIC(10,2),
  match_status          TEXT,                             -- "Verified (paystub)" / "Pending Verification"
  sent_status           TEXT DEFAULT 'No',
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_payroll_org_week ON payroll_entries(org_id, week_start);

-- Invoices & AR (mirrors 15-column Invoices & AR sheet)
CREATE TABLE invoices (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  invoice_number        TEXT NOT NULL,
  invoice_date          DATE NOT NULL,
  due_date              DATE,
  contractor_id         UUID REFERENCES contractors(id),
  contract_num          TEXT,
  region_code           TEXT,
  wo_id                 UUID REFERENCES work_orders(id),
  description           TEXT,
  sqft                  NUMERIC(12,2),
  rate                  NUMERIC(10,4),
  amount                NUMERIC(12,2),
  status                TEXT NOT NULL DEFAULT 'draft'
                          CHECK (status IN ('draft','sent','paid')),
  qb_invoice_id         TEXT,
  qb_doc_number         TEXT,
  payment_received      BOOLEAN NOT NULL DEFAULT false,
  payment_date          DATE,
  notes                 TEXT,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE(org_id, invoice_number)
);

-- Photos
CREATE TABLE photos (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  wo_id                 UUID NOT NULL REFERENCES work_orders(id) ON DELETE CASCADE,
  storage_key           TEXT NOT NULL,
  filename              TEXT NOT NULL,
  mime_type             TEXT NOT NULL,
  size_bytes            BIGINT,
  thumbnail_key         TEXT,
  latitude              NUMERIC(10,7),
  longitude             NUMERIC(10,7),
  address               TEXT,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_photos_wo ON photos(wo_id);

-- Signatures (stored separately from photos)
CREATE TABLE signatures (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  wo_id                 UUID REFERENCES work_orders(id),
  document_id           UUID REFERENCES documents(id),
  storage_key           TEXT NOT NULL,
  filename              TEXT NOT NULL,
  signer_name           TEXT,
  signer_title          TEXT,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Audit Log (mirrors 8-column Automation Log)
CREATE TABLE audit_log (
  id                    BIGSERIAL PRIMARY KEY,
  org_id                UUID NOT NULL REFERENCES organizations(id),
  user_id               UUID REFERENCES users(id),
  source                TEXT,                             -- "Scan Inbox", "Invoice Generator", etc.
  action                TEXT NOT NULL,
  subject               TEXT,                             -- "WO PT-12345", "Doc PL_2026-08-03"
  details               JSONB,
  status                TEXT,                             -- "Detected", "Error", "Generated"
  action_required       TEXT,                             -- "Yes — Enter WO data" or null
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_audit_org_date ON audit_log(org_id, created_at);
```

### Integration & Job Tables

```sql
-- Per-tenant integration connections
CREATE TABLE integrations (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  type                  TEXT NOT NULL,                    -- "google_drive", "quickbooks"
  enabled               BOOLEAN NOT NULL DEFAULT false,
  config                JSONB NOT NULL DEFAULT '{}',      -- connector-specific config
  credentials           TEXT,                             -- encrypted OAuth tokens / API keys
  status                TEXT NOT NULL DEFAULT 'disconnected'
                          CHECK (status IN ('disconnected','connected','error')),
  last_sync_at          TIMESTAMPTZ,
  error_message         TEXT,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE(org_id, type)
);

-- Job queue tracking (mirrors BullMQ state for UI visibility)
CREATE TABLE jobs (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id                UUID NOT NULL REFERENCES organizations(id),
  type                  TEXT NOT NULL,                    -- "fill_production_log", "scan_wo", etc.
  status                TEXT NOT NULL DEFAULT 'pending'
                          CHECK (status IN ('pending','processing','completed','failed')),
  payload               JSONB NOT NULL,                   -- input data (same JSON structure as current Drive JSONs)
  result                JSONB,                            -- output: {storage_key, filename, doc_id, etc.}
  error                 TEXT,
  attempts              INT NOT NULL DEFAULT 0,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  completed_at          TIMESTAMPTZ
);

CREATE INDEX idx_jobs_org_status ON jobs(org_id, status);
CREATE INDEX idx_jobs_type_status ON jobs(type, status);
```

---

## Integration Interfaces

### Storage Interface

```typescript
interface StorageConnector {
  upload(orgId: string, path: string, data: Buffer, mime: string): Promise<string>
  download(orgId: string, storageKey: string): Promise<Buffer>
  getSignedUrl(orgId: string, storageKey: string, expiresInSec: number): Promise<string>
  delete(orgId: string, storageKey: string): Promise<void>
  list(orgId: string, prefix: string): Promise<FileMetadata[]>
}

// Default: R2Connector (always available, no tenant config needed)
// Optional: GoogleDriveConnector (one-way push: after storing in R2, copy to tenant's Drive)
```

### Email Interface

```typescript
interface EmailConnector {
  send(orgId: string, params: {
    to: string | string[],
    subject: string,
    html: string,
    attachments?: { filename: string, content: Buffer }[]
  }): Promise<void>
}

// Default: ResendConnector (platform account)
```

### Accounting Interface (Designed Now, Built Later)

```typescript
interface AccountingConnector {
  createInvoice(orgId: string, invoice: InvoiceData): Promise<{ externalId: string, docNumber: string }>
  getInvoice(orgId: string, externalId: string): Promise<InvoiceData | null>
  lookupCustomer(orgId: string, name: string): Promise<{ customerId: string } | null>
  getInvoiceUrl(orgId: string, externalId: string): Promise<string>
}

// Future: QuickBooksConnector (port existing qb.js behind this interface)
```

### Geocoding Interface

```typescript
interface GeocodingConnector {
  geocode(address: string, options?: { bounds?: LatLngBounds }): Promise<GeoResult>
  reverseGeocode(lat: number, lng: number): Promise<AddressResult>
}

// Default: GoogleMapsConnector (platform API key)
```

---

## Job Queue Design

### Job Types

```typescript
type JobType =
  | 'fill_production_log'
  | 'fill_certified_payroll'
  | 'fill_signin'
  | 'fill_field_report'
  | 'fill_month_end'
  | 'scan_work_order'
  | 'process_approved_docs'
  | 'sync_to_drive'
  | 'send_email'
```

### Flow (Replaces Drive Polling)

```
1. API creates job → Redis (BullMQ)
2. Python worker consumes job:
   - Downloads template from R2 (cached locally)
   - Fills PDF using existing fill_*.py code (UNCHANGED)
   - Uploads result to R2
   - Updates documents table via API callback or direct DB
3. If tenant has Drive integration: enqueues 'sync_to_drive' job
4. If doc type has email recipients: enqueues 'send_email' job
5. Status visible via documents table (real-time, no polling folders)
```

### Python Worker Changes (Minimal)

- Replace `watch_and_fill.py` Drive poll loop → Redis/BullMQ consumer
- Replace Drive upload → R2 upload
- Replace Apps Script proxy calls → direct DB or API calls
- **Fill logic is UNCHANGED**: `fill_production_log.py`, `fill_certified_payroll.py`, etc.

---

## Auth & Tenant Middleware

```typescript
// Every request:
// 1. Extract JWT from HttpOnly cookie
// 2. Verify → { userId, orgId, role }
// 3. Attach to request context

function authMiddleware(req, res, next) {
  const token = req.cookies['session']
  if (!token) return res.status(401).json({ error: 'Not authenticated' })
  const payload = verifyJwt(token)
  req.user = { id: payload.userId, orgId: payload.orgId, role: payload.role }
  next()
}

function requireRole(...allowed: Role[]) {
  return (req, res, next) => {
    if (!allowed.includes(req.user.role))
      return res.status(403).json({ error: 'Forbidden' })
    next()
  }
}

// EVERY database query: WHERE org_id = req.user.orgId
// No exceptions. This is the tenant isolation boundary.
```

### Role Permissions

| Capability | Owner | Admin | Foreman | Crew |
|-----------|-------|-------|---------|------|
| Organization settings | ✓ | | | |
| User management | ✓ | ✓ | | |
| All WOs + financial data | ✓ | ✓ | | |
| Assigned WOs only | | | ✓ | ✓ |
| Document approval | ✓ | ✓ | | |
| Generate documents | ✓ | ✓ | ✓ | |
| Field report entry | ✓ | ✓ | ✓ | ✓ |
| Sign-in entry | ✓ | ✓ | ✓ | ✓ |
| Revenue/production dashboards | ✓ | ✓ | | |
| Connect integrations | ✓ | ✓ | | |

---

## API Surface

### Auth
```
POST   /api/auth/signup
POST   /api/auth/login
POST   /api/auth/logout
POST   /api/auth/refresh
GET    /api/auth/me
```

### Tenant Configuration
```
GET/PATCH  /api/settings/org
CRUD       /api/settings/regions
CRUD       /api/settings/contractors
CRUD       /api/settings/categories
CRUD       /api/settings/pricing
CRUD       /api/settings/classifications
CRUD       /api/settings/pay-rates
GET/PATCH  /api/settings/overtime
CRUD       /api/settings/employees
CRUD       /api/settings/users
POST       /api/settings/users/invite
CRUD       /api/settings/billing-remaps
CRUD       /api/settings/contract-lookup
```

### Work Orders
```
GET    /api/wos                       — List (filtered, paginated)
POST   /api/wos                       — Create (from scan or manual)
GET    /api/wos/:id                   — Single WO with computed rollups
PATCH  /api/wos/:id                   — Update (status, coordinates, etc.)
DELETE /api/wos/:id                   — Delete (cascades marking items)
GET    /api/wos/map                   — WOs with coordinates for map view
POST   /api/wos/:id/waterblast/confirm — Confirm waterblasting
```

### Marking Items
```
GET    /api/wos/:woId/markings        — List items for WO
POST   /api/wos/:woId/markings        — Create item
PATCH  /api/markings/:id              — Update item
DELETE /api/markings                   — Bulk delete
```

### Sign-In
```
GET    /api/signin/queue              — Pending entries
POST   /api/signin                    — Submit sign-in
POST   /api/signin/check-continuation — Shift continuation check
POST   /api/signin/day-hours          — Hours summary for date
GET    /api/signin/rows/:fileId       — Rows for approval editing
GET    /api/signin/header/:fileId     — Header for approval display
POST   /api/signin/rows/:fileId/edit  — Edit hours in approval
```

### Field Reports
```
POST   /api/field-reports             — Submit
POST   /api/field-reports/check-shift — Validate shift attribution
POST   /api/field-reports/finalize    — Batch finalize
```

### Documents
```
GET    /api/documents/pending         — Approval queue
GET    /api/documents/pending/counts  — Badge counts
GET    /api/documents/:id/pdf         — Stream PDF
GET    /api/documents/:id/meta        — Metadata
POST   /api/documents/:id/approve     — Approve
POST   /api/documents/:id/approve-with-signature — Approve + overlay
POST   /api/documents/:id/skip-signoff — Approve without signature
POST   /api/documents/:id/regenerate  — Regenerate from current data
POST   /api/documents/:id/reupload    — Replace PDF
GET    /api/documents/status          — Doc status calendar
POST   /api/documents/status/flags    — Toggle done/sent
POST   /api/documents/batch/list      — List for batch download
POST   /api/documents/batch/download  — Stream ZIP
POST   /api/documents/flags           — WO-level doc flags (CFR done/sent, invoice done/sent)
```

### Photos
```
POST   /api/wos/:woId/photos          — Upload
GET    /api/wos/:woId/photos          — List with thumbnails
GET    /api/photos/:id/content        — Full-res
DELETE /api/photos/:id                — Delete
```

### Tools / Generation
```
POST   /api/tools/daily-documents     — Generate PL/SI/CFR for date
POST   /api/tools/certified-payroll   — Generate CP for week
POST   /api/tools/month-end           — Generate EU/Certificates
POST   /api/tools/month-end/all       — Batch month-end → ZIP
POST   /api/tools/process-approved    — Archive approved docs
POST   /api/tools/scan-wo             — Upload + parse WO scan
GET    /api/tools/scan-status         — Poll scan results
GET    /api/tools/scan-uploads-today  — Today's scans
POST   /api/tools/paystub/parse       — Parse paystub image
```

### Dashboards
```
GET    /api/dashboard                 — WO list + stats
GET    /api/revenue                   — Revenue data (date range)
GET    /api/production                — Production data (date range)
GET    /api/pending-counts            — Nav badge counts
GET    /api/pending-counts/doc-status — Doc status badge
```

### Geocoding
```
POST   /api/geocode                   — Forward geocode
POST   /api/reverse-geocode           — Reverse geocode
```

### Integrations
```
GET    /api/integrations              — List available + connected
POST   /api/integrations/:type/connect — Start connection flow
POST   /api/integrations/:type/disconnect — Remove connection
GET    /api/integrations/:type/status — Health check
```

### QuickBooks (Deferred — Interface Ready)
```
GET    /api/qb/auth-start             — Start OAuth
GET    /api/qb/auth-callback          — OAuth callback
GET    /api/qb/status                 — Connection status
GET    /api/qb/disconnect             — Disconnect webhook
POST   /api/qb/invoice/:woId         — Generate invoice
```

---

## Feature Parity Checklist

| Feature | Day One | Deferred |
|---------|---------|----------|
| WO Tracker + Dashboard | ✓ | |
| Field Reports + Marking Items | ✓ | |
| Sign-In + Time Tracking | ✓ | |
| Document Generation (PL, CP, CFR, SI) | ✓ | |
| Document Approval Workflow | ✓ | |
| Revenue Dashboard | ✓ | |
| Production Dashboard | ✓ | |
| WO Scanning (Claude Vision) | ✓ | |
| Nav/Map Tab | ✓ | |
| Doc Status Calendar | ✓ | |
| Batch Document Download | ✓ | |
| Month-End Docs (EU, Certificates) | ✓ | |
| Multi-Crew Support | ✓ | |
| Photo Pipeline (watermark, upload) | ✓ | |
| QuickBooks Invoicing | | ✓ |
| Google Drive Sync | | ✓ (post-launch) |

---

## Repository Structure

The new system lives under `platform/` alongside the existing code. The current system (`Code.js`, `workers/`, `webapp/`) stays untouched and keeps running on Railway until migration is complete. PDF templates in `templates/` are shared by both systems.

```
oneiro-ops-automation/
│
│   ═══════════════ CURRENT SYSTEM (untouched, keeps running) ═══════════════
│
├── Code.js                              # Apps Script — deployed via clasp
├── appsscript.json                      # Apps Script manifest
├── .claspignore
├── workers/                             # Python workers — deployed on Railway
│   ├── watch_and_fill.py                #   Drive poll loop (replaced by platform/worker)
│   ├── fill_production_log.py
│   ├── fill_certified_payroll.py
│   ├── fill_signin.py
│   ├── fill_contractor_field_report.py
│   ├── fill_utilization.py
│   ├── fill_form.py
│   ├── fill_server.py
│   ├── _appearances.py
│   ├── parse_work_order.py
│   └── __init__.py
├── webapp/                              # React + Express — deployed on Railway
│   ├── server.js
│   ├── server/
│   │   ├── qb.js
│   │   └── qbItems.js
│   ├── src/
│   ├── public/
│   ├── package.json
│   ├── vite.config.js
│   ├── tailwind.config.js
│   ├── postcss.config.js
│   ├── index.html
│   └── .env.example
├── fill_production_log.py               # Root-level dead code (legacy local-only)
├── fill_certified_payroll.py            # Root-level dead code (legacy local-only)
├── watch_and_fill.py                    # Root-level dead code (legacy local-only)
├── Procfile                             # Current Railway: worker + web
├── requirements.txt                     # Current Python deps
├── .python-version
├── .gitignore
├── Mail_Audit_Fillable.pdf
│
│   ═══════════════ SHARED ═══════════════
│
├── templates/                           # PDF templates — used by BOTH systems
│   ├── Metro_Production_Log_Fillable.pdf
│   ├── Certified_Payroll_Fillable.pdf
│   ├── Sign_In_Log_Fillable.pdf
│   ├── Thermo_Contractor_Field_Report_Fillable.pdf
│   └── Monthly_Workforce_Utilization_Fillable.pdf
│
├── docs/                                # Documentation
│   ├── pricing_engine.md
│   ├── marking_type_mapping.md
│   └── mvp_design_plan.md              # This file
│
│   ═══════════════ NEW SYSTEM ═══════════════
│
└── platform/
    │
    ├── docker-compose.yml               # Local dev: postgres + redis
    ├── Procfile                          # Railway: api + worker processes
    ├── README.md
    │
    ├── api/                             # ── TypeScript/Node API Server ──
    │   ├── package.json
    │   ├── tsconfig.json
    │   ├── .env.example
    │   │
    │   └── src/
    │       ├── index.ts                 # Entry point — starts Express, connects DB
    │       ├── app.ts                   # Express app setup, middleware chain, static serving
    │       ├── config.ts                # Typed env var config (DB, Redis, R2, Resend, Maps, Claude)
    │       │
    │       ├── middleware/
    │       │   ├── auth.ts              # JWT extraction from cookie → req.user {id, orgId, role}
    │       │   ├── tenant.ts            # Ensures org_id is present, loads org config
    │       │   └── roles.ts             # requireRole('owner','admin',...) guard
    │       │
    │       ├── routes/
    │       │   ├── auth.ts              # POST signup/login/logout/refresh, GET me
    │       │   ├── settings/
    │       │   │   ├── org.ts           # GET/PATCH organization
    │       │   │   ├── regions.ts       # CRUD regions
    │       │   │   ├── contractors.ts   # CRUD contractors (incl. doc distribution flags)
    │       │   │   ├── categories.ts    # CRUD marking categories
    │       │   │   ├── pricing.ts       # CRUD contract pricing + pricing multipliers
    │       │   │   ├── payroll.ts       # CRUD classifications, pay rates, OT rules
    │       │   │   ├── employees.ts     # CRUD employee registry
    │       │   │   ├── users.ts         # CRUD users + POST invite
    │       │   │   ├── billingRemaps.ts # CRUD billing remap rules
    │       │   │   └── contractLookup.ts # CRUD contract lookup (contract ID / reg #)
    │       │   ├── workOrders.ts        # CRUD WOs, map view, waterblast confirm
    │       │   ├── markingItems.ts      # CRUD marking items (nested under WO)
    │       │   ├── signin.ts            # Queue, submit, continuation check, day-hours, row editing
    │       │   ├── fieldReports.ts      # Submit, shift attribution check, finalize
    │       │   ├── documents.ts         # Pending queue, approve (all variants), regenerate,
    │       │   │                        #   reupload, status calendar, done/sent flags,
    │       │   │                        #   batch list, batch download (ZIP), PDF streaming
    │       │   ├── photos.ts            # Upload, list, content stream, delete
    │       │   ├── tools.ts             # Generate daily docs, cert payroll, month-end,
    │       │   │                        #   process approved, scan WO, scan status, paystub parse
    │       │   ├── dashboards.ts        # Dashboard data, revenue, production, pending counts
    │       │   ├── geocoding.ts         # Forward + reverse geocode
    │       │   └── integrations.ts      # List, connect, disconnect, status
    │       │
    │       ├── services/                # ── Business Logic (SINGLE SOURCE OF TRUTH) ──
    │       │   ├── pricing.ts           # Pricing engine: rate resolution, group multipliers,
    │       │   │                        #   revenue calculation (replaces Code.js + webapp/pricing.js)
    │       │   ├── overtime.ts          # OT calculation: daily/weekly thresholds, weekend rules,
    │       │   │                        #   cross-group lookback (replaces Code.js + signinShared.js)
    │       │   ├── markingAggregation.ts # Production log aggregation: category → PL row mapping,
    │       │   │                        #   per-crew-chief filtering, SF/paint rollups
    │       │   ├── payroll.ts           # Certified payroll: week bucketing, rate resolution,
    │       │   │                        #   supplemental calculation, gross/net computation
    │       │   ├── revenue.ts           # Revenue dashboard: daily/by-contractor/by-group aggregation,
    │       │   │                        #   needs-pricing detection, labor cost computation
    │       │   ├── production.ts        # Production dashboard: SF/LF/EA by day/contractor,
    │       │   │                        #   shift KPIs, streak tracking
    │       │   ├── docLifecycle.ts      # Document lifecycle: status transitions, done/sent flags,
    │       │   │                        #   doc key generation, dedup, calendar computation
    │       │   ├── docGeneration.ts     # Document generation orchestration: daily doc triggers,
    │       │   │                        #   CP triggers, month-end spec building, job creation
    │       │   ├── billingRemap.ts      # Billing remap: rule resolution, date-gated cutover,
    │       │   │                        #   month-granularity variant for EU/CERT
    │       │   ├── opDay.ts             # Operational day: timestamp bucketing using org cutoff hour
    │       │   ├── woScanning.ts        # WO scan orchestration: job creation, status polling,
    │       │   │                        #   result processing into WO + marking items
    │       │   └── pdfOverlay.ts        # PDF signature overlay via pdf-lib (used during approval)
    │       │
    │       ├── integrations/            # ── Connector Interfaces + Implementations ──
    │       │   ├── storage/
    │       │   │   ├── interface.ts      # StorageConnector interface
    │       │   │   ├── r2.ts            # Cloudflare R2 implementation (default)
    │       │   │   └── googleDrive.ts   # Google Drive one-way push (optional, deferred)
    │       │   ├── email/
    │       │   │   ├── interface.ts      # EmailConnector interface
    │       │   │   └── resend.ts        # Resend implementation (default)
    │       │   ├── geocoding/
    │       │   │   ├── interface.ts      # GeocodingConnector interface
    │       │   │   └── googleMaps.ts    # Google Maps implementation (default)
    │       │   └── accounting/
    │       │       ├── interface.ts      # AccountingConnector interface
    │       │       └── quickbooks.ts    # QuickBooks implementation (deferred)
    │       │
    │       ├── db/
    │       │   ├── client.ts            # Drizzle client + connection pool
    │       │   ├── schema.ts            # Drizzle table definitions (all tables)
    │       │   └── queries/             # Query functions organized by domain
    │       │       ├── organizations.ts
    │       │       ├── users.ts
    │       │       ├── workOrders.ts
    │       │       ├── markingItems.ts
    │       │       ├── signin.ts
    │       │       ├── documents.ts
    │       │       ├── payroll.ts
    │       │       ├── invoices.ts
    │       │       ├── photos.ts
    │       │       ├── settings.ts      # Regions, contractors, categories, pricing, pay rates, etc.
    │       │       └── audit.ts
    │       │
    │       ├── jobs/
    │       │   ├── queue.ts             # BullMQ queue setup + connection
    │       │   ├── producers.ts         # Helper functions to enqueue each job type
    │       │   └── types.ts             # Job type enum + payload type definitions
    │       │
    │       └── utils/
    │           ├── jwt.ts               # Sign/verify JWT, cookie helpers
    │           ├── password.ts          # bcrypt hash/compare
    │           ├── encryption.ts        # AES-256-GCM for sensitive fields (tax_id, ssn, tokens)
    │           ├── validation.ts        # Zod schemas for request body validation
    │           ├── money.ts             # Scaled-integer arithmetic (matches QB precision)
    │           └── errors.ts            # AppError class, error handler middleware
    │
    ├── worker/                          # ── Python PDF Worker ──
    │   ├── main.py                      # Job consumer: connects to Redis, routes jobs to fillers
    │   ├── storage.py                   # R2 upload/download (replaces Drive API calls)
    │   ├── db.py                        # Direct DB connection for status updates (or API callback)
    │   ├── requirements.txt             # Python deps (same + redis client)
    │   │
    │   │   # ── Fill modules (copied from workers/, adapted I/O) ──
    │   ├── fill_production_log.py       # Multi-page PL fill (same logic, R2 I/O)
    │   ├── fill_certified_payroll.py    # Multi-page CP fill (same logic, R2 I/O)
    │   ├── fill_signin.py              # Sign-in fill + signature overlay (same logic)
    │   ├── fill_contractor_field_report.py  # CFR fill + WO merge via Vision (same logic)
    │   ├── fill_utilization.py          # Monthly EU form fill (same logic)
    │   ├── fill_form.py                 # Generic AcroForm filler (unchanged)
    │   ├── fill_server.py               # Month-end HTTP fill endpoint (same, R2 I/O)
    │   ├── _appearances.py              # PyMuPDF appearance regeneration (unchanged)
    │   └── parse_work_order.py          # Claude Vision WO parser (unchanged)
    │
    ├── web/                             # ── React Frontend (adapted from webapp/) ──
    │   ├── package.json
    │   ├── vite.config.ts
    │   ├── tailwind.config.js
    │   ├── postcss.config.js
    │   ├── index.html                   # Same meta tags, favicon
    │   │
    │   ├── public/
    │   │   ├── vendor/
    │   │   │   ├── opencv.js            # Self-hosted OpenCV.js (~9MB, lazy-loaded)
    │   │   │   └── jscanify.min.js      # Corner detection for doc scanning
    │   │   ├── legal/
    │   │   │   ├── privacy.html
    │   │   │   └── eula.html
    │   │   └── favicon.svg
    │   │
    │   └── src/
    │       ├── main.jsx                 # React DOM render
    │       ├── index.css                # Tailwind directives + custom classes (card, btn-*, etc.)
    │       ├── App.jsx                  # Router: auth guard, nav, route table
    │       │
    │       ├── pages/                   # ── Existing pages (same UI, adapted API calls) ──
    │       │   ├── Dashboard.jsx        # WO Tracker table, stat cards, charts, tools menu
    │       │   ├── FieldReport.jsx      # Marking items form (top table, grid, manual)
    │       │   ├── SignIn.jsx           # Crew time tracking, signatures, upload/scan
    │       │   ├── Approvals.jsx        # PDF viewer, approve/sign/regenerate
    │       │   ├── ScanWO.jsx           # WO upload, ready queue, history
    │       │   ├── RevenueTab.jsx       # Revenue KPIs, charts, pricing group breakdown
    │       │   ├── ProductionTab.jsx    # SF/LF/EA metrics, contractor breakdown
    │       │   ├── DocStatusTab.jsx     # Calendar view, day/week popovers, done/sent toggles
    │       │   ├── NavTab.jsx           # Google Maps with WO pins, geocode editing
    │       │   │
    │       │   │                        # ── New pages ──
    │       │   ├── Login.jsx            # Email + password login
    │       │   ├── Signup.jsx           # Create account + organization
    │       │   ├── AcceptInvite.jsx     # Accept email invitation, set password
    │       │   └── Settings.jsx         # Organization config: sub-pages for org info,
    │       │                            #   regions, contractors, categories, pricing,
    │       │                            #   payroll, employees, users, integrations
    │       │
    │       ├── components/              # ── Existing components (unchanged) ──
    │       │   ├── ConfirmModal.jsx
    │       │   ├── DeleteWOModal.jsx
    │       │   ├── DocStatusChips.jsx
    │       │   ├── DownloadDocumentsModal.jsx
    │       │   ├── EditCoordinatesModal.jsx
    │       │   ├── FilterBar.jsx
    │       │   ├── GenerateDocModal.jsx
    │       │   ├── InvoiceCell.jsx
    │       │   ├── MarkingFormModal.jsx
    │       │   ├── PaystubUpload.jsx
    │       │   ├── PrincipalSignModal.jsx
    │       │   ├── QBStatusBadge.jsx
    │       │   ├── ReuploadModal.jsx
    │       │   ├── RowKebab.jsx
    │       │   ├── ScanCapture.jsx
    │       │   ├── SignInHeaderCard.jsx
    │       │   ├── SignInHoursEditor.jsx
    │       │   ├── SignaturePad.jsx
    │       │   ├── StatusBadge.jsx
    │       │   ├── StatusPickerModal.jsx
    │       │   └── WODocsQueue.jsx
    │       │
    │       └── lib/                     # ── Existing lib modules (unchanged) ──
    │           ├── PendingCountsContext.jsx  # Shared nav badge state
    │           ├── dateOps.js           # Operational day helpers
    │           ├── docScanner.js        # OpenCV + jscanify wrapper
    │           ├── imagesToPdf.js       # Multi-image → PDF conversion
    │           ├── markingCategories.js # Category definitions, layout logic
    │           ├── parseQty.js          # Arithmetic quantity parsing (15x10 → 150)
    │           ├── photoPipeline.js     # Photo watermarking pipeline
    │           ├── photoUploadQueue.js  # IndexedDB queue + concurrent uploader
    │           ├── pricing.js           # Client-side pricing preview
    │           ├── qtyValidation.js     # Quantity range warnings
    │           └── signinShared.js      # Client-side OT preview
    │
    ├── migrations/                      # ── Database Migrations ──
    │   └── 0001_initial_schema.sql      # Full schema from design plan
    │
    └── seed/                            # ── Default Data for New Tenants ──
        ├── nyc_dot_categories.json      # 48 marking categories (name, unit, pricing_group, form_section, etc.)
        ├── nyc_dot_multipliers.json     # Line width multipliers, extruded/preformed unit counts
        ├── nyc_dot_classifications.json # LP, SAT, OP, LAB, FGL, SUP + default pay rates
        └── nyc_regions.json             # MN, BK, BX, QU, SI with full names
```

### Design Decisions Reflected in Structure

**Why `platform/` and not a separate repo:**
The old code (`Code.js`, `workers/`, `webapp/`) serves as living reference documentation during the build. Claude and the team can cross-reference the existing 15,174-line Code.js while implementing each service. Once migration is complete, the old code can be archived to a `legacy/` folder.

**Why copy Python fill modules instead of importing:**
The adapted worker will have different I/O (R2 instead of Drive, Redis queue instead of Drive polling). Having a clean copy means the old `workers/` directory can eventually be deleted without breaking the new system. The fill logic itself (~3,500 lines) is small enough that copying is cleaner than sharing.

**Why copy React source instead of symlinking:**
The frontend needs real changes (auth pages, settings pages, removal of access gate code, removal of QB UI for now). Copying gives a clean starting point without risk of breaking the running production app.

**Why `routes/settings/` is a subdirectory:**
There are 10 settings-related route files. Keeping them flat under `routes/` would clutter the directory. Grouping them under `settings/` makes the structure scannable.

**Why `services/` contains both pure logic and orchestration:**
- Pure business logic: `pricing.ts`, `overtime.ts`, `opDay.ts`, `billingRemap.ts`, `money.ts`
- Orchestration: `docGeneration.ts`, `woScanning.ts`, `docLifecycle.ts`
- PDF manipulation: `pdfOverlay.ts` (signature overlay via pdf-lib during approval)

These all live under `services/` because they're the backend's domain layer — the routes call them, they call the DB and integrations.

**Why `db/queries/` instead of a full repository pattern:**
Query files are organized by domain (workOrders, markingItems, signin, etc.) but they're just functions that take the Drizzle client and return typed results. No abstract repository classes. This keeps things simple and lets Claude produce focused query functions that can be tested independently.

**Why `utils/money.ts`:**
The existing system uses scaled-integer arithmetic (`_money2_` in Code.js) to match QuickBooks' exact decimal precision. This utility must exist in the new system to prevent 1-cent rounding errors on invoices.

**Why `utils/encryption.ts`:**
Fields marked "encrypted at app level" in the schema (tax_id, ssn_last4, OAuth tokens) need a shared encryption utility. AES-256-GCM, same approach as the current QB token encryption in `qb.js`.

### What's NOT in the New System (Intentionally)

- No `server/qb.js` or `server/qbItems.js` — deferred. The `integrations/accounting/quickbooks.ts` file exists as interface-only.
- No Google Drive sync — deferred. The `integrations/storage/googleDrive.ts` file exists as a placeholder.
- No `QBStatusBadge.jsx` functionality — component copied but will be inactive until QB connector ships.
- No `InvoiceCell.jsx` QB link functionality — same.

### Railway Deployment (New System)

The `platform/Procfile` defines two processes:

```
web: node api/dist/index.js
worker: python worker/main.py
```

Railway deploys from `platform/` directory. The API server serves the built React SPA from `web/dist/` in production mode. Both services share the same Railway environment variables for DB, Redis, R2, and API keys.
