/**
 * Pricing engine — SINGLE SOURCE OF TRUTH.
 *
 * Ported from Code.js priceMarkingItem_ (lines 10024–10101),
 * _resolveRateRow_ (lines 9740–9767), and _money2_ (lines 9994–9999).
 *
 * Also replaces webapp/src/lib/pricing.js (client-side preview copy).
 *
 * Revenue formula per pricing group:
 *   line4:         qty × rate_line4         × LINE_WIDTH_MULTIPLIER[category]
 *   line12:        qty × rate_line12        × LINE12_MULTIPLIER[category]
 *   preformed:     qty × rate_preformed     × PREFORMED_UNIT_COUNT[category]
 *   extruded:      qty × rate_extruded      × EXTRUDED_UNIT_COUNT[category]
 *   color_surface: qty × rate_color_surface (no multiplier)
 *   unpriced:      always $0, surfaces in "Needs Pricing" bucket
 */

// ─── Multiplier Tables ──────────────────────────────────────

/** Line width → multiplier of 4" line base rate. */
export const LINE_WIDTH_MULTIPLIER: Record<string, number> = {
  '4" Line':            1.0,
  '6" Line':            1.5,
  '8" Line':            2.0,
  '12" Line':           3.0,
  '16" Line':           4.0,
  '24" Line':           6.0,
  'Lane Lines':         10,    // EA (10' skips) → qty × base × 10
  'Double Yellow Line': 2.0,
  'Double White Line':  2.0,
};

/** HVX Crosswalk = base 12" rate; Stop Line = 24" stripe at 2×. */
export const LINE12_MULTIPLIER: Record<string, number> = {
  'HVX Crosswalk': 1.0,
  'Stop Line':     2.0,
};

/**
 * Extruded unit counts — from NYC DOT 8' Letters & Numbers table.
 * Message values are sums of per-letter unit counts:
 *   STOP = S(0.37)+T(0.25)+O(0.39)+P(0.34) = 1.35
 *   ONLY = O(0.39)+N(0.46)+L(0.25)+Y(0.25) = 1.35
 *   BUS  = B(0.46)+U(0.36)+S(0.37)         = 1.19
 *   BUMP = B(0.46)+U(0.36)+M(0.48)+P(0.34) = 1.64
 *   20MPH= 2(0.37)+0(0.39)+M(0.48)+P(0.34)+H(0.39) = 1.97
 *   RR   = R(0.41)+R(0.41) = 0.82
 *   X    = X(0.31) = 0.31
 */
export const EXTRUDED_UNIT_COUNT: Record<string, number> = {
  'Stop Msg':               1.35,
  'Only Msg':               1.35,
  'Bus Msg':                1.19,
  'Bump Msg':               1.64,
  '20 MPH Msg':             1.97,
  'Railroad (RR)':          0.82,
  'Railroad (X)':           0.31,
  'L/R Arrow':              1.00,
  'Straight Arrow':         0.81,
  'Combination Arrow':      1.65,
  'Combination Arrow (L/R)': 1.74,
  'Speed Hump Markings':    0.78,
  'Shark Teeth 12x18':      0.05,
  'Shark Teeth 24x36':      0.19,
  'Bike Lane Arrow':        0.29,
};

/** Preformed symbol unit counts. */
export const PREFORMED_UNIT_COUNT: Record<string, number> = {
  'Bike Lane Symbol':            0.91,  // legacy alias for Old Bike Symbol
  'Old Bike Symbol (w/ rider)':  0.91,
  'New Bike Symbol (just bike)': 0.97,
  'Pedestrian Men':              0.84,
};

/** Revenue bucket definitions for dashboard KPIs. */
export const REVENUE_BUCKETS = [
  { key: 'thermo',  label: 'Thermo',  groups: ['line4', 'line12', 'extruded'] },
  { key: 'mma',     label: 'MMA',     groups: ['color_surface'] },
  { key: 'preform', label: 'Preform', groups: ['preformed'] },
] as const;

/** Human-readable reasons why an item can't be priced. */
export const NEEDS_PRICING_REASON_LABEL: Record<string, string> = {
  no_rate:           'No Contract Pricing row matches this contract',
  no_unit_count:     'Unit count missing from table',
  unpriced_category: 'Category requires manual pricing',
  unit_migration:    'Bike Lane Green Bar entered as EA — re-enter as SF',
  no_date:           'Date Completed is blank',
  bad_qty:           'Quantity is blank, zero, or not a number',
};

// ─── Types ───────────────────────────────────────────────────

export type PricingGroup = 'line4' | 'line12' | 'preformed' | 'extruded' | 'color_surface' | 'unpriced';

export interface RateRow {
  contractor: string;
  contractNum: string;
  regionCode: string;
  effectiveDate: Date | null;
  rates: {
    line4: number | null;
    line12: number | null;
    preformed: number | null;
    extruded: number | null;
    colorSurface: number | null;
  };
}

export interface PricingResult {
  revenue: number;
  group: PricingGroup;
  reason: string | null;  // null = priced cleanly
  rate?: number;          // $/unit actually applied (4dp)
}

export interface MarkingItemForPricing {
  category: string;
  quantity: number | string | null;
  unit?: string | null;
  dateCompleted?: string | null;
}

export interface WoMeta {
  contractor: string;
  contractNum: string;
  regionCode: string;
}

// ─── Core Functions ──────────────────────────────────────────

/**
 * Resolve the pricing group for a marking category.
 * Reads from the tenant's marking_categories table (pricingGroup field).
 * Falls back to 'unpriced' if not found.
 */
export function resolvePricingGroup(
  category: string,
  categoryMap: Record<string, string>,
): PricingGroup {
  return (categoryMap[category] || 'unpriced') as PricingGroup;
}

/**
 * Pick the rate row that applies to a given (contractor, contract, region)
 * on the supplied date. Exact port of Code.js _resolveRateRow_.
 *
 * Resolution order:
 * 1. Dated rows whose effectiveDate <= itemDate, newest first
 * 2. Blank-date rows (effective forever)
 * 3. null if nothing matches
 */
export function resolveRateRow(
  rates: RateRow[],
  contractor: string,
  contractNum: string,
  regionCode: string,
  dateIso: string | null,
): RateRow | null {
  if (!rates.length) return null;
  const cn = String(contractNum || '').split('/')[0].trim();

  const itemDate = dateIso ? new Date(dateIso) : null;
  if (itemDate && isNaN(itemDate.getTime())) return null;

  const candidates = rates.filter(r =>
    r.contractor === contractor &&
    r.contractNum === cn &&
    r.regionCode === regionCode,
  );
  if (candidates.length === 0) return null;

  // Prefer dated rows whose effectiveDate <= item date
  const dated = candidates.filter(r => r.effectiveDate !== null);
  const applicable = (itemDate
    ? dated.filter(r => r.effectiveDate!.getTime() <= itemDate.getTime())
    : dated
  ).sort((a, b) => b.effectiveDate!.getTime() - a.effectiveDate!.getTime());

  if (applicable.length > 0) return applicable[0];

  // Fall back to blank-date row
  return candidates.find(r => r.effectiveDate === null) ?? null;
}

/**
 * Compute revenue for one marking item. Exact port of Code.js priceMarkingItem_.
 *
 * @param item - The marking item (category, quantity, unit, dateCompleted)
 * @param woMeta - Work order metadata (contractor, contractNum, regionCode)
 * @param rates - All rate rows for this org (from contract_pricing table)
 * @param categoryGroupMap - Mapping of category name → pricing group (from marking_categories table)
 * @param multipliers - Tenant's multiplier overrides (from pricing_multipliers table)
 */
export function priceMarkingItem(
  item: MarkingItemForPricing,
  woMeta: WoMeta,
  rates: RateRow[],
  categoryGroupMap: Record<string, string>,
  multipliers: {
    lineWidth: Record<string, number>;
    line12: Record<string, number>;
    extrudedUnit: Record<string, number>;
    preformedUnit: Record<string, number>;
  },
): PricingResult {
  const cat = String(item.category || '').trim();
  if (!cat) {
    return { revenue: 0, group: 'unpriced', reason: 'unpriced_category' };
  }

  const qty = Number(item.quantity);
  if (isNaN(qty) || qty <= 0) {
    return { revenue: 0, group: 'unpriced', reason: 'bad_qty' };
  }

  // Bike Lane Green Bar legacy unit migration check
  if (cat === 'Bike Lane Green Bar' && String(item.unit || '').toUpperCase() === 'EA') {
    return { revenue: 0, group: 'color_surface', reason: 'unit_migration' };
  }

  const group = resolvePricingGroup(cat, categoryGroupMap);
  if (group === 'unpriced') {
    return { revenue: 0, group, reason: 'unpriced_category' };
  }

  if (!item.dateCompleted) {
    return { revenue: 0, group, reason: 'no_date' };
  }

  const rateRow = resolveRateRow(rates, woMeta.contractor, woMeta.contractNum, woMeta.regionCode, item.dateCompleted);
  if (!rateRow) {
    return { revenue: 0, group, reason: 'no_rate' };
  }

  switch (group) {
    case 'line4': {
      const base = rateRow.rates.line4;
      if (base === null) return { revenue: 0, group, reason: 'no_rate' };
      const mult = multipliers.lineWidth[cat];
      if (mult === undefined) return { revenue: 0, group, reason: 'no_unit_count' };
      const r = rate4(base * mult);
      return { revenue: money2(qty, r), group, reason: null, rate: r };
    }
    case 'line12': {
      const base = rateRow.rates.line12;
      if (base === null) return { revenue: 0, group, reason: 'no_rate' };
      const mult = multipliers.line12[cat];
      if (mult === undefined) return { revenue: 0, group, reason: 'no_unit_count' };
      const r = rate4(base * mult);
      return { revenue: money2(qty, r), group, reason: null, rate: r };
    }
    case 'preformed': {
      const base = rateRow.rates.preformed;
      if (base === null) return { revenue: 0, group, reason: 'no_rate' };
      const units = multipliers.preformedUnit[cat];
      if (units === undefined) return { revenue: 0, group, reason: 'no_unit_count' };
      const r = rate4(base * units);
      return { revenue: money2(qty, r), group, reason: null, rate: r };
    }
    case 'extruded': {
      const base = rateRow.rates.extruded;
      if (base === null) return { revenue: 0, group, reason: 'no_rate' };
      const units = multipliers.extrudedUnit[cat];
      if (units === undefined) return { revenue: 0, group, reason: 'no_unit_count' };
      const r = rate4(base * units);
      return { revenue: money2(qty, r), group, reason: null, rate: r };
    }
    case 'color_surface': {
      const base = rateRow.rates.colorSurface;
      if (base === null) return { revenue: 0, group, reason: 'no_rate' };
      const r = rate4(base);
      return { revenue: money2(qty, r), group, reason: null, rate: r };
    }
    default:
      return { revenue: 0, group: 'unpriced', reason: 'unpriced_category' };
  }
}

// ─── Arithmetic Helpers ──────────────────────────────────────

/**
 * Round a rate to 4 decimal places.
 * Exact port of Code.js _rate4_ (line 9974).
 */
export function rate4(n: number): number {
  return Math.round(Number(n) * 10000) / 10000;
}

/**
 * Qty × Rate → Revenue (2 decimal places).
 * Exact port of Code.js _money2_ (lines 9994–9999).
 *
 * Uses scaled-integer arithmetic to avoid floating-point errors.
 * Converts qty to hundredths and rate to ten-thousandths, multiplies
 * (producing ten-millionths of a dollar), half-rounds up to cents.
 */
export function money2(qty: number, rate: number): number {
  const q = Math.round(Number(qty) * 100);       // hundredths
  const r = Math.round(Number(rate) * 10000);     // ten-thousandths
  const micro = q * r;                            // 1e-6 dollars, exact
  return Math.floor((micro + 5000) / 10000) / 100; // half-up to cents
}
