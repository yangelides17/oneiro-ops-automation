# API Behavioral Gap Audit — Old Apps Script vs New TypeScript API

Generated 2026-08-26. Exhaustive line-by-line comparison across all flows.

---

## CRITICAL GAPS (must fix — broken behavior or data loss)

### Field Report (11 gaps)
| # | Gap | Detail |
|---|-----|--------|
| FR-1 | No operational day correction | Uses UTC wall-clock date instead of opToday(). Night crew submitting at 3 AM gets wrong date for WDL, marking items, workEndDate |
| FR-2 | Status state machine missing | Only ever sets 'completed' or leaves unchanged. Never writes Dispatched or advances Received→Dispatched→In Progress |
| FR-3 | dispatchDate / workStartDate never auto-set | Old code auto-derives both from first FR's date if blank. New code only sets workEndDate on complete |
| FR-4 | Issues overwritten, not appended | Old: appends with date prefix. New: `issues \|\| wo.issuesReported` — last-write-wins, destroys history |
| FR-5 | finalizeMarkingStatus not called on non-complete submits | Partial-day work (crew fills qtys but doesn't mark WO complete) never commits items. Qty stays pending indefinitely |
| FR-6 | crewChief not required | Old returns 400 if missing. New accepts optional — breaks sign-in queue grouping |
| FR-7 | MMA waterblast gate missing server-side | Old blocks submit if waterblast not confirmed on MMA WOs. New has no check |
| FR-8 | CFR doc key returns '' — document creation silently skipped | `'field_report'` not in DOC_TYPE_PREFIX → `buildDocKey()` returns '' → entire CFR block never runs |
| FR-9 | CFR fill job payload missing ~12 fields | school, from/to streets, install dates, markings dict, grid, prep_by, etc. all absent |
| FR-10 | WDL entry uses UTC date, not operational day | Same root cause as FR-1 but separate impact on sign-in queue grouping |
| FR-11 | finalize endpoint has wrong semantics | Old: generates CFR JSON. New: approves existing docs. Completely different concepts |

### Marking Items (5 gaps)
| # | Gap | Detail |
|---|-----|--------|
| MI-1 | No server-side unit auto-derivation from category | Old enforces CATEGORY_UNITS on create and update. New accepts whatever client sends |
| MI-2 | qty=0 PATCH does not revert status to Pending | Old explicitly reverts and clears dateCompleted. New leaves Completed+qty=0 |
| MI-3 | Stop Msg/Lines not expanded per-direction | Old expands "EW" into 2 rows (E + W). New creates 1 row with no direction |
| MI-4 | HVX detection requires literal "HVX" string | Old creates row for any non-empty grid cell value. New requires `.includes('HVX')` |
| MI-5 | computeMarkingRollups has no equivalent | marking_types, quantity_completed, paint_material never computed or stored |

### Work Orders (5 gaps)
| # | Gap | Detail |
|---|-----|--------|
| WO-1 | WO list sort by creation date, not status priority | Active WOs (In Progress) don't surface at top of dropdown |
| WO-2 | DELETE will FK-violate on Work Day Log | WDL has RESTRICT FK to work_orders — delete fails if any WDL rows exist |
| WO-3 | edit_completed_wo endpoint entirely missing | Admin rework path (batched edits + CFR regen modes) not ported |
| WO-4 | POST /api/wos does not seed marking items | Old seeded from top_markings/intersection_grid. Only scan path seeds |
| WO-5 | PT- WO auto-seed of Color Surface item missing | Old auto-adds a Bike Lane (SF) default item for Paint/MMA WOs |

### Sign-In (10 gaps)
| # | Gap | Detail |
|---|-----|--------|
| SI-1 | Work Day Log never marked "Submitted" | Queue cards never clear after sign-in submission |
| SI-2 | No billing remap on submit | Sign-in rows stored under raw contract/region instead of billing identity |
| SI-3 | wo_ids array → single woId per row | Multi-WO sign-in model structurally lost |
| SI-4 | No sign-in document/JSON generation | No filled sign-in PDFs produced. No file written for filler worker |
| SI-5 | No upload path (source='uploaded') | Manual PDF upload flow not implemented |
| SI-6 | Sign-in queue sort order reversed | New: newest first. Old: oldest first (FIFO — surfaces overdue work) |
| SI-7 | Queue response missing 7 fields | bill_contract_number, bill_borough, prime_contractor, subcontractor, address, contract_id, project_name |
| SI-8 | Check-continuation logic completely replaced | Old: 60-min gap detection across dates. New: just counts existing rows |
| SI-9 | Day-hours scope narrowed + response shape changed | Old: all groups on date → `{totals: {name: hours}}`. New: single group → `{hours: [{employeeName, totalHours, totalOt}]}` |
| SI-10 | Sign-in header for doc missing 7 fields | bill_contract/borough, contractor, prime_contractor, subcontractor, address, wos array |

### Documents (3 gaps)
| # | Gap | Detail |
|---|-----|--------|
| DOC-1 | Approve/skip-signoff: no file movement or archive trigger | Old moves Drive file to Approved folder for cron pickup. New only sets DB flag |
| DOC-2 | Doc status calendar response completely flat | Old: rich per-day/contractor/week structure with color rollup. New: flat doc list |
| DOC-3 | No per-doc PL/CP generate endpoints | generate_pl_for_doc and generate_cp_for_doc have no equivalent |

### Dashboard (2 gaps)
| # | Gap | Detail |
|---|-----|--------|
| DASH-1 | Dashboard WOs missing docs{} object + invoice fields + folder_url | All doc lifecycle chips show "not done". Invoice column always blank |
| DASH-2 | Revenue top_wos hardcoded [] | "Top Work Orders" table always empty |

### Tools (3 gaps)
| # | Gap | Detail |
|---|-----|--------|
| TOOL-1 | Daily doc generation missing SI completeness gate | Old validates all sign-ins done before generating. New generates regardless |
| TOOL-2 | Certified payroll generation missing SI completeness gate | Same as TOOL-1 |
| TOOL-3 | Month-end generation is a stub | No fill spec, field data, or group structure computed |

### Frontend (6 gaps)
| # | Gap | Detail |
|---|-----|--------|
| FE-1 | Status values snake_case vs Title Case | DB: 'in_progress'. Frontend filters: 'In Progress'. Every filter/badge breaks |
| FE-2 | Sign-in submit body sends old field names | Frontend sends {queue_id, wo_ids, contractor...}. API expects {rows: [{woId, contractorId...}]}. Every submit 400s |
| FE-3 | Day-hours response `{totals:{}}` vs `{hours:[]}` | Shift Totals "Other" column always shows 0 |
| FE-4 | Photo gallery completely broken | No thumbnail_b64, no drive_file_id, no mime. Lightbox, delete, thumbnails all fail |
| FE-5 | Scan-status poll: POST to GET endpoint + field mismatch | Poll always 404s (wrong HTTP method). Response uses jobId not fileId |
| FE-6 | scan-uploads-today response shape mismatch | Missing fileId, woIds, is_combined, uploaded_at. All committed items render empty |

### Settings (2 gaps)
| # | Gap | Detail |
|---|-----|--------|
| SET-1 | Overtime rules PATCH returns 404 on fresh orgs | No POST to create initial row. Admin can never set rules for new org |
| SET-2 | Pricing create always fails — UI sends empty contractorId | UUID validation rejects '' |

### Other (3 gaps)
| # | Gap | Detail |
|---|-----|--------|
| OTH-1 | FieldReport deep-link ?wo=RM-xxx never matches | .id is UUID not WO number |
| OTH-2 | waterBlastRequired_required typo in FieldReport.jsx | Waterblast gate always bypassed in UI |
| OTH-3 | No idempotency on sign-in submit | Network retry creates duplicate rows |

---

## MEDIUM GAPS (degraded but functional)

| # | Area | Gap |
|---|------|-----|
| M-1 | WO | Missing audit log entry on PATCH status change |
| M-2 | WO | WO create skips Contract ID auto-lookup from Contract Lookup table |
| M-3 | WO | WO list missing folder_url field |
| M-4 | MI | Grid category validation (intersection+direction required) missing |
| M-5 | MI | Bike lane description format "Per WO: {qty}" → "Source: {source}" |
| M-6 | MI | Manual create allows status override (old always forced Pending) |
| M-7 | FR | photos_uploaded sticky flag not tracked |
| M-8 | FR | check-shift overnight attribution detection not ported |
| M-9 | SI | Queue entry missing subcontractor, address, project_name |
| M-10 | SI | Time format: 24h stored vs 12h in old system |
| M-11 | SI | Sign-in row edits: no ambiguity guard for multi-chief |
| M-12 | DOC | approved_docs_pending hardcoded 0 |
| M-13 | DOC | Batch download missing counts/missing/warnings breakdown |
| M-14 | DOC | Payroll-period batch endpoint doesn't exist |
| M-15 | DOC | Regenerate has no pre-flight validation (SI completeness) |
| M-16 | DASH | Stats key 'complete' vs 'completed'; attention missing photos check |
| M-17 | DASH | Revenue labor breakdown loses MMA/thermo split |
| M-18 | DASH | Revenue needs_pricing items missing contractor/contract fields |
| M-19 | DASH | Revenue by_group not zero-filled for missing groups |
| M-20 | DASH | Production top_wos sorted by item count, not total quantity |
| M-21 | DASH | Production by_contractor field: 'contractor' → 'contractorName' |
| M-22 | DASH | Revenue pct_invoiced camelCase vs snake_case |
| M-23 | DASH | Pending counts: signins_pending absent, doc_status_pending logic differs |
| M-24 | SET | Contractor delete returns raw 500 FK error instead of 409 |
| M-25 | SET | Contractor document delivery flags have no defaults |
| M-26 | SET | Employee DELETE missing; Settings.jsx delete calls 404 |
| M-27 | SET | /api/employees returns inactive employees (no isActive filter) |
| M-28 | SET | Role update accepts any arbitrary string |
| M-29 | SET | No endpoint to revoke invitations or remove users |
| M-30 | SET | Category delete can orphan marking items |
| M-31 | SET | Classification delete can orphan pay rates |
| M-32 | SET | Dashboard WOs missing invoice fields |
| M-33 | SET | FieldReport WO panel missing folderUrl |
| M-34 | TOOL | process-approved response missing errored/skipped fields |
| M-35 | PHOTO | Photo upload response shape changed — frontend can't map IDs |
| M-36 | GEO | Reverse-geocode response shape needs verification |
| M-37 | QB | QB record/clear/refresh-token/customer-id endpoints missing (not just stubs) |

---

## LOW GAPS (cosmetic or minor)

| # | Area | Gap |
|---|------|-----|
| L-1 | WO | Delete response missing marking_items_deleted/wdl_preserved counts |
| L-2 | WO | PATCH noop-on-same-status check missing |
| L-3 | MI | Skipped items with qty could be inadvertently completed (edge case) |
| L-4 | MI | N+1 query loop on finalize (functional but slow) |
| L-5 | FR | Automation Log operational notes less detailed |
| L-6 | SI | Audit log detail reduced vs old Automation Log |
| L-7 | SI | OT rewrite is unconditional (extra writes, not incorrect) |
| L-8 | SI | Classification validation looser (accepts any string) |
| L-9 | DOC | Reupload orphans old R2 file (not deleted) |
| L-10 | DASH | No caching (performance, not correctness) |
| L-11 | DASH | Revenue default range uses UTC instead of TZ |
| L-12 | DASH | Production pct_days_worked rounded to int vs 1 decimal |
| L-13 | DASH | Production longest_streak DST drift risk |
| L-14 | SET | Region sortOrder has no default seeding |
| L-15 | SET | formSection field unused by frontend |
| L-16 | SET | suppSt/suppOt no guidance or seeding |
| L-17 | SET | Billing remap has no PATCH (acceptable delete+recreate) |
| L-18 | SET | Contract Lookup upsert silently overwrites |
| L-19 | SET | Employee SSN/sensitive fields returned to browser on settings load |
| L-20 | TOOL | Daily doc response date_used → date rename (frontend doesn't use it) |

---

## TOTALS

| Severity | Count |
|----------|-------|
| CRITICAL | 50 |
| MEDIUM | 37 |
| LOW | 20 |
| **TOTAL** | **107** |
