/**
 * Billing remap service — SINGLE SOURCE OF TRUTH.
 *
 * Ported from Code.js _billingRemap_ (lines 9879–9886),
 * _billingRemapAsOf_ (lines 9902–9909),
 * _billingRemapForMonth_ (lines 9917–9924),
 * _hasBillingRemap_ (lines 9932–9936).
 *
 * In the old system, remap rules were hardcoded in _BILLING_REMAP_.
 * In the new system, they come from the billing_remaps database table.
 *
 * Purpose: When a sub-prime does work on a contract they didn't win,
 * the billing identity (for documents, invoices, pricing) differs from
 * the raw operational identity. This service resolves that mapping.
 */

export interface RemapRule {
  sourceContract: string;
  sourceRegion: string;
  sourceContractor: string | null;  // null = matches any contractor
  targetContract: string;
  targetRegion: string;
  effectiveDate: string;  // YYYY-MM-DD — rule applies from this date on
}

export interface BillingIdentity {
  contractNum: string;
  regionCode: string;
}

/**
 * Always-on base remap. Exact port of Code.js _billingRemap_.
 *
 * Looks up the first matching rule for (contractNum, regionCode, contractor).
 * Returns the billing identity if a rule matches, or the raw identity if not.
 */
export function billingRemap(
  rules: RemapRule[],
  contractNum: string,
  regionCode: string,
  contractor: string,
): BillingIdentity {
  const cn = String(contractNum || '').trim();
  const br = String(regionCode || '').trim();
  const co = String(contractor || '').trim();

  const hit = rules.find(r =>
    r.sourceContract === cn &&
    r.sourceRegion === br &&
    (r.sourceContractor === null || r.sourceContractor === co),
  );

  return hit
    ? { contractNum: hit.targetContract, regionCode: hit.targetRegion }
    : { contractNum: cn, regionCode: br };
}

/**
 * Date-gated remap. Exact port of Code.js _billingRemapAsOf_.
 *
 * Before a rule's effective date: returns raw identity.
 * From effective date onward: applies the remap.
 * Strips '/EXT' suffix from contractNum in both branches.
 *
 * @param dateIso - YYYY-MM-DD date string (the work/anchor date)
 */
export function billingRemapAsOf(
  rules: RemapRule[],
  dateIso: string | null,
  contractNum: string,
  regionCode: string,
  contractor: string,
): BillingIdentity {
  const cn = String(contractNum || '').split('/')[0].trim();
  const br = String(regionCode || '').trim();

  if (!dateIso) {
    return { contractNum: cn, regionCode: br };
  }

  // Filter to rules whose effective date <= the given date
  const applicableRules = rules.filter(r =>
    String(dateIso) >= r.effectiveDate,
  );

  if (applicableRules.length === 0) {
    return { contractNum: cn, regionCode: br };
  }

  return billingRemap(applicableRules, cn, br, contractor);
}

/**
 * Month-granularity remap. Exact port of Code.js _billingRemapForMonth_.
 *
 * The entire effective month rolls up to billing — otherwise days before
 * the effective date within the same month would key raw and split the
 * month's identity in two. Months before the effective month stay raw.
 *
 * @param monthIso - YYYY-MM month string
 */
export function billingRemapForMonth(
  rules: RemapRule[],
  monthIso: string | null,
  contractNum: string,
  regionCode: string,
  contractor: string,
): BillingIdentity {
  const cn = String(contractNum || '').split('/')[0].trim();
  const br = String(regionCode || '').trim();

  if (!monthIso) {
    return { contractNum: cn, regionCode: br };
  }

  // Filter to rules whose effective month <= the given month
  const applicableRules = rules.filter(r =>
    String(monthIso) >= r.effectiveDate.slice(0, 7),
  );

  if (applicableRules.length === 0) {
    return { contractNum: cn, regionCode: br };
  }

  return billingRemap(applicableRules, cn, br, contractor);
}

/**
 * Check if any remap rule targets the given (contractNum, regionCode).
 * Exact port of Code.js _hasBillingRemap_.
 *
 * Used by CP generator to decide whether to split a bucket by contractor.
 */
export function hasBillingRemap(
  rules: RemapRule[],
  contractNum: string,
  regionCode: string,
): boolean {
  const cn = String(contractNum || '').trim();
  const br = String(regionCode || '').trim();
  return rules.some(r => r.sourceContract === cn && r.sourceRegion === br);
}
