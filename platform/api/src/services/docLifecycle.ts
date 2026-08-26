/**
 * Document lifecycle service — SINGLE SOURCE OF TRUTH.
 *
 * Ported from Code.js _docLifecycleId_ (lines 7093–7106),
 * _DOC_TYPE_PREFIX_ (lines 7088–7092), MONTH_END_DOCS_ (lines 7128–7131),
 * and _upsertDocLifecycleRow_ / _setDocLifecycleStatus_ logic.
 *
 * Manages document identity (key generation), status transitions,
 * and done/sent flag tracking.
 */

// ─── Doc Key Generation ──────────────────────────────────────

/** Maps doc type names to key prefixes (SI, CP, CFR — PL has its own function). */
const DOC_TYPE_PREFIX: Record<string, string> = {
  'signin':            'SI',
  'certified_payroll': 'CP',
  'field_report':      'CFR',
};

/** Month-end document definitions. */
export const MONTH_END_DOCS = [
  { key: 'EU',   docType: 'employee_utilization', label: 'Employee Utilization' },
  { key: 'CERT', docType: 'certificates',         label: 'Certificates' },
] as const;

/**
 * Generate a Production Log document key.
 * Exact port of Code.js _plDocId_ (line 7945).
 *
 * PL keys use CONTRACTOR NAME (not contract+region) because PLs are
 * per-contractor, not per-contract-region like SI/CP.
 *
 * Format: PL_{anchorDate}_{contractorSlug}[_chief-{chiefSlug}]
 * Examples:
 *   PL_2026-08-24_Metro_Express
 *   PL_2026-08-24_Metro_Express_chief-BobSmith
 */
export function buildPlDocKey(
  anchorDate: string,
  contractor: string,
  crewChief?: string | null,
): string {
  const slug = String(contractor || '').trim().replace(/\s+/g, '_');
  const base = `PL_${String(anchorDate).trim()}_${slug}`;
  const chiefSlug = String(crewChief || '').replace(/[^A-Za-z0-9]/g, '');
  return chiefSlug ? `${base}_chief-${chiefSlug}` : base;
}

/**
 * Generate a Sign-In or Certified Payroll document key.
 * Exact port of Code.js _docLifecycleId_ (lines 7093–7106).
 *
 * SI/CP keys use CONTRACT+REGION (not contractor name).
 *
 * Format: {PREFIX}_{anchorDate}_{contractNum}_{regionCode}[_chief-{slug}]
 * Examples:
 *   SI_2026-08-24_84125MBTP701_BK
 *   SI_2026-08-24_84125MBTP701_BK_chief-JohnSmith
 *   CP_2026-08-18_84125MBTP701_BK
 */
export function buildDocKey(
  docType: string,
  anchorDate: string,
  contractNum: string,
  regionCode: string,
  crewChief?: string | null,
): string {
  const prefix = DOC_TYPE_PREFIX[docType];
  if (!prefix) return '';

  const cn = String(contractNum || '').split('/')[0].trim();
  const base = `${prefix}_${String(anchorDate).trim()}_${cn}_${String(regionCode).trim()}`;

  const chiefSlug = String(crewChief || '').replace(/[^A-Za-z0-9]/g, '');
  return chiefSlug ? `${base}_chief-${chiefSlug}` : base;
}

/**
 * Generate a month-end document key.
 * Format: {KEY}_{YYYY-MM}_{contractNum}_{regionCode}_{contractorSlug}
 * Examples:
 *   EU_2026-07_84125MBTP701_BK_MetroExpress
 *   CERT_2026-07_84125MBTP701_BK_MetroExpress
 */
export function buildMonthEndDocKey(
  monthEndKey: string, // 'EU' or 'CERT'
  monthIso: string,    // 'YYYY-MM'
  contractNum: string,
  regionCode: string,
  contractor: string,
): string {
  const cn = String(contractNum || '').split('/')[0].trim();
  const slug = String(contractor || '').replace(/[^A-Za-z0-9]/g, '');
  return `${monthEndKey}_${monthIso}_${cn}_${String(regionCode).trim()}_${slug}`;
}

/**
 * Parse a SI/CP doc key back into its components.
 * Returns null if the key doesn't match expected format.
 */
export function parseDocKey(docKey: string): {
  prefix: string;
  anchorDate: string;
  contractNum: string;
  regionCode: string;
  crewChief?: string;
} | null {
  // SI/CP format: PREFIX_YYYY-MM-DD_CONTRACT_REGION[_chief-SLUG]
  const stdMatch = docKey.match(/^(SI|CP)_(\d{4}-\d{2}-\d{2})_([^_]+)_([^_]+?)(?:_chief-(.+))?$/);
  if (stdMatch) {
    return {
      prefix: stdMatch[1],
      anchorDate: stdMatch[2],
      contractNum: stdMatch[3],
      regionCode: stdMatch[4],
      crewChief: stdMatch[5] || undefined,
    };
  }
  return null;
}

/**
 * Parse a PL doc key back into its components.
 * PL keys use contractor slug, not contract+region.
 */
export function parsePlDocKey(docKey: string): {
  anchorDate: string;
  contractorSlug: string;
  crewChief?: string;
} | null {
  const plMatch = docKey.match(/^PL_(\d{4}-\d{2}-\d{2})_(.+?)(?:_chief-(.+))?$/);
  if (plMatch) {
    return {
      anchorDate: plMatch[1],
      contractorSlug: plMatch[2],
      crewChief: plMatch[3] || undefined,
    };
  }
  return null;
}

/**
 * Parse a month-end doc key.
 */
export function parseMonthEndDocKey(docKey: string): {
  key: string;
  monthIso: string;
  contractNum: string;
  regionCode: string;
  contractorSlug: string;
} | null {
  const m = docKey.match(/^(EU|CERT)_(\d{4}-\d{2})_([^_]+)_([^_]+)_(.+)$/);
  if (!m) return null;
  return {
    key: m[1],
    monthIso: m[2],
    contractNum: m[3],
    regionCode: m[4],
    contractorSlug: m[5],
  };
}

// ─── Status Transitions ─────────────────────────────────────

export type DocStatus = 'pending' | 'needs_review' | 'approved' | 'archived';

/** Valid status transitions. */
const VALID_TRANSITIONS: Record<DocStatus, DocStatus[]> = {
  pending:      ['needs_review'],
  needs_review: ['approved', 'pending'],  // can be sent back for re-generation
  approved:     ['archived'],
  archived:     [],                       // terminal state
};

/**
 * Check if a status transition is valid.
 */
export function isValidTransition(from: DocStatus, to: DocStatus): boolean {
  return VALID_TRANSITIONS[from]?.includes(to) ?? false;
}

// ─── Payroll Week Helpers ────────────────────────────────────

/**
 * Get the Sunday that starts the payroll week containing a given date.
 * Payroll weeks run Sunday–Saturday.
 */
export function payrollWeekStart(dateIso: string): string {
  const m = dateIso.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return dateIso;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const dow = d.getDay(); // 0=Sun
  d.setDate(d.getDate() - dow);
  return formatDate(d);
}

/**
 * Get the Saturday that ends the payroll week.
 */
export function payrollWeekEnd(weekStartIso: string): string {
  const m = weekStartIso.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return weekStartIso;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  d.setDate(d.getDate() + 6);
  return formatDate(d);
}

/**
 * Get the month (YYYY-MM) for month-end documents.
 * Uses the payroll week's Saturday to determine the month — if a week
 * straddles month boundaries, it belongs to the month its Saturday falls in.
 */
export function payrollMonthIso(dateIso: string): string {
  const ws = payrollWeekStart(dateIso);
  const we = payrollWeekEnd(ws);
  return we.slice(0, 7);
}

function formatDate(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}
