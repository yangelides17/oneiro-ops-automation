/**
 * Marking Item Business Logic
 *
 * Server-side enforcement of marking item rules that the old Apps Script
 * embedded inline in handleCreateMarkingItem_, handleUpdateMarkingItem_,
 * finalizeMarkingStatus_, seedMarkingItems_, and expandDirLetters_.
 *
 * Ported from Code.js:
 *   CATEGORY_UNITS_          (lines 9478–9529)
 *   unitForCategory_         (line  9531)
 *   expandDirLetters_        (lines 10104–10120)
 *   finalizeMarkingStatus_   (lines 10868–10927)
 *   computeMarkingRollups_   (lines 10430–10486)
 *   handleCreateMarkingItem_ (lines 10633–10705)  — validation subset
 *   handleUpdateMarkingItem_ (lines 10729–10812)  — qty-clearing logic
 */

import { eq, and, inArray } from 'drizzle-orm';
import type { Db } from '../db/client.js';
import { markingItems } from '../db/schema.js';

// ─── Category → Unit Map ──────────────────────────────────────
// Single source of truth. Mirrors the frontend's CATEGORY_UNITS
// (web/src/lib/markingCategories.js) and Code.js CATEGORY_UNITS_
// (line 9478). A category present in this map has a LOCKED unit —
// the server enforces it regardless of what the client sends.
// "Others" is intentionally absent (variable unit, user picks).

export const CATEGORY_UNITS: Record<string, string> = {
  // SF — MMA area work
  'Bike Lane':           'SF',
  'Bus Lane':            'SF',
  'Pedestrian Space':    'SF',

  // LF — lines, crosswalks, stop lines
  'Double Yellow Line':  'LF',
  'Double White Line':   'LF',
  'Lane Lines':          'EA', // 10' skips — each EA is one 10 LF skip of 4" line
  'Solid Lines':         'LF',
  '4" Line':             'LF',
  '6" Line':             'LF',
  '8" Line':             'LF',
  '12" Line':            'LF',
  '16" Line':            'LF',
  '24" Line':            'LF',
  'Gores':               'LF',
  'HVX Crosswalk':       'LF',
  'Stop Line':           'LF',

  // EA — messages, arrows, misc
  'Messages':            'EA',
  'Stop Msg':            'EA',
  'Only Msg':            'EA',
  'Bus Msg':             'EA',
  'Bump Msg':            'EA',
  'Custom Msg':          'EA',
  '20 MPH Msg':          'EA',
  'Railroad (RR)':       'EA',
  'Railroad (X)':        'EA',
  'Rail Road X/Diamond': 'EA',
  'Arrows':              'EA',
  'L/R Arrow':           'EA',
  'Straight Arrow':      'EA',
  'Combination Arrow':   'EA',
  'Combination Arrow (L/R)': 'EA',
  'Speed Hump Markings': 'EA',
  'Shark Teeth 12x18':   'EA',
  'Shark Teeth 24x36':   'EA',
  'Bike Lane Arrow':     'EA',
  'Bike Lane Symbol':    'EA', // legacy — retired from picker, kept for old rows
  'Old Bike Symbol (w/ rider)':  'EA',
  'New Bike Symbol (just bike)': 'EA',
  'Pedestrian Men':      'EA',
  'Bike Lane Green Bar': 'SF',
};

/** Categories that require intersection + direction on create. */
const GRID_CATEGORIES = new Set(['HVX Crosswalk', 'Stop Line', 'Stop Msg']);

// ─── Pure Functions ───────────────────────────────────────────

/**
 * Returns the locked unit for a category, or '' if variable.
 * Exact port of Code.js unitForCategory_ (line 9531).
 */
export function deriveUnit(category: string): string {
  return CATEGORY_UNITS[String(category || '').trim()] || '';
}

/**
 * Returns true if the category has a fixed unit (not user-pickable).
 */
export function isUnitLocked(category: string): boolean {
  return deriveUnit(category) !== '';
}

/**
 * Validate that grid categories have required intersection + direction.
 * Exact port of Code.js handleCreateMarkingItem_ (lines 10643–10648).
 */
export function validateGridCategory(
  category: string,
  intersection?: string | null,
  direction?: string | null,
): string | null {
  if (!GRID_CATEGORIES.has(category)) return null;
  if (!String(intersection || '').trim()) {
    return `Intersection is required for ${category}.`;
  }
  if (!String(direction || '').trim()) {
    return `Direction is required for ${category}.`;
  }
  return null;
}

/**
 * Expand a directional string into individual direction letters.
 * Exact port of Code.js expandDirLetters_ (lines 10104–10120).
 *
 * Examples:
 *   "North"  → ["N"]
 *   "EW"     → ["E", "W"]
 *   "NSEW"   → ["N", "S", "E", "W"]
 *   ""       → []
 */
export function expandDirections(val: string): string[] {
  const s = String(val || '').trim();
  if (!s) return [];

  const FULL_WORD: Record<string, string> = {
    NORTH: 'N', EAST: 'E', SOUTH: 'S', WEST: 'W',
  };
  const upper = s.toUpperCase();

  // Full word match (e.g. "North" → ["N"])
  if (FULL_WORD[upper]) return [FULL_WORD[upper]];

  // Treat as concatenated letters, dedupe preserving order
  const seen = new Set<string>();
  const out: string[] = [];
  for (const c of upper) {
    if (!'NSEW'.includes(c)) continue;
    if (seen.has(c)) continue;
    seen.add(c);
    out.push(c);
  }
  return out;
}

/**
 * Enforce create defaults: always Pending status, derive unit from
 * category, default addedBy to 'manual'.
 *
 * Port of Code.js handleCreateMarkingItem_ (lines 10677–10695).
 * Returns the cleaned data ready for DB insertion.
 */
export function enforceCreateDefaults(data: {
  category: string;
  quantity?: number | null;
  unit?: string | null;
  status?: string;
  addedBy?: string;
  woSection?: string;
  [key: string]: unknown;
}): typeof data & { unit: string; status: string; addedBy: string; woSection: string } {
  const lockedUnit = deriveUnit(data.category);
  return {
    ...data,
    unit: lockedUnit || String(data.unit || 'EA').trim(),
    status: 'pending',           // always Pending on create, never client-overridable
    addedBy: data.addedBy || 'manual',
    woSection: data.woSection || 'manual',
  };
}

/**
 * Apply update-specific business rules to a marking item patch.
 *
 * Rules (exact port of Code.js handleUpdateMarkingItem_, lines 10772–10806):
 *
 * 1. If category changes → auto-derive unit from CATEGORY_UNITS.
 *    If the new category has a locked unit, it overrides whatever
 *    the client sent. Variable categories (e.g. "Others") honor
 *    the client-supplied unit.
 *
 * 2. If quantity is being set to 0/null/empty → revert status to
 *    'pending' and clear dateCompleted. This is the explicit
 *    "didn't get done" signal.
 *
 * 3. Changing quantity to a different POSITIVE value leaves status
 *    and dateCompleted intact (same-day correction).
 *
 * Returns the augmented patch with any derived fields added.
 */
export function applyUpdateRules(
  patch: Record<string, unknown>,
  existingCategory?: string,
): Record<string, unknown> {
  const result = { ...patch };

  // Rule 1: category change → re-derive unit
  if (result.category !== undefined) {
    const newCat = String(result.category || '').trim();
    const lockedUnit = deriveUnit(newCat);
    if (lockedUnit) {
      result.unit = lockedUnit;
    }
  }

  // Rule 2: quantity clearing → revert to pending
  if (result.quantity !== undefined) {
    const q = typeof result.quantity === 'number'
      ? result.quantity
      : parseFloat(String(result.quantity));
    const hasQty = !isNaN(q) && q > 0;

    if (!hasQty) {
      result.status = 'pending';
      result.dateCompleted = null;
    }
    // Rule 3: positive qty change — leave status/dateCompleted intact
    // (no action needed, they stay as-is)
  }

  // If category didn't change but unit might need enforcing on the
  // existing category (e.g. client sends unit='EA' for an HVX Crosswalk)
  if (result.unit !== undefined && result.category === undefined && existingCategory) {
    const lockedUnit = deriveUnit(existingCategory);
    if (lockedUnit) {
      result.unit = lockedUnit;
    }
  }

  return result;
}

// ─── Database Operations ──────────────────────────────────────

/**
 * Finalize marking item status for a WO.
 * Exact port of Code.js finalizeMarkingStatus_ (lines 10868–10927).
 *
 * Called on EVERY field report submit (not just wo_complete=yes).
 * This is how partial-day work gets committed: crew fills in qtys,
 * submits, items with qty>0 become Completed with that day's date.
 *
 * Rules:
 *   - qty > 0 AND status != 'completed' → promote to completed,
 *     set dateCompleted, tag crewChief
 *   - qty > 0 AND already completed → leave untouched (preserves
 *     Day 1 date across Day 2 resubmits)
 *   - qty empty/0 → force pending, clear dateCompleted
 *   - status = 'skipped' → leave untouched regardless of qty
 */
export async function finalizeMarkingStatus(
  db: Db,
  orgId: string,
  woId: string,
  dateOfWork: string,
  crewChief?: string,
): Promise<{ promoted: number; reverted: number }> {
  const items = await db.select()
    .from(markingItems)
    .where(and(eq(markingItems.orgId, orgId), eq(markingItems.woId, woId)));

  let promoted = 0;
  let reverted = 0;
  const chief = String(crewChief || '').trim();

  for (const item of items) {
    // Skip items explicitly marked as skipped by admin/foreman
    if (item.status === 'skipped') continue;

    const qty = parseFloat(String(item.quantity || ''));
    const hasQty = !isNaN(qty) && qty > 0;

    if (hasQty && item.status !== 'completed') {
      // Promote: Pending → Completed with this day's date
      const updates: Record<string, unknown> = {
        status: 'completed',
        dateCompleted: dateOfWork,
        updatedAt: new Date(),
      };
      if (chief) updates.crewChief = chief;

      await db.update(markingItems)
        .set(updates)
        .where(eq(markingItems.id, item.id));
      promoted++;
    } else if (!hasQty && item.status !== 'pending') {
      // Revert: qty empty/0 cannot be Completed — force back to Pending
      await db.update(markingItems)
        .set({
          status: 'pending',
          dateCompleted: null,
          updatedAt: new Date(),
        })
        .where(eq(markingItems.id, item.id));
      reverted++;
    }
    // hasQty && already completed → no-op (preserve original date)
  }

  return { promoted, reverted };
}

// ─── Marking Rollups ──────────────────────────────────────────

export interface MarkingRollups {
  markingTypes: string;
  quantityCompleted: number | null;
  paintMaterial: string;
}

/**
 * Compute WO-level marking rollup columns.
 * Exact port of Code.js computeMarkingRollups_ (lines 10430–10486).
 *
 * Rules:
 *   - Thermo WOs: marking_types = 'N/A', paint_material = 'N/A',
 *     quantity_completed = SUM(qty WHERE unit='LF')
 *   - MMA WOs: marking_types = distinct completed categories,
 *     paint_material = distinct color/material values,
 *     quantity_completed = SUM(qty WHERE unit='SF')
 *   - EA-unit items are always excluded from the quantity sum
 */
export function computeMarkingRollups(
  items: {
    workType?: string | null;
    category?: string | null;
    quantity?: string | null;
    unit?: string | null;
    colorMaterial?: string | null;
    status?: string | null;
  }[],
  woWorkType?: string,
): MarkingRollups {
  const blank: MarkingRollups = { markingTypes: '', quantityCompleted: null, paintMaterial: '' };
  if (items.length === 0) return blank;

  // Determine if this is a thermo or MMA WO.
  // Port of Code.js line 10444: checks per-item workType first (any item
  // marked thermo → thermo path), then falls back to woWorkType.
  const anyItemThermo = items.some(r =>
    String(r.workType || '').toLowerCase() === 'thermo'
  );
  const isThermo = anyItemThermo
    || (woWorkType?.toLowerCase() === 'thermo')
    || (!woWorkType && !items.some(r => r.workType));

  const targetUnit = isThermo ? 'LF' : 'SF';

  let qtySum = 0;
  let hasQty = false;

  for (const item of items) {
    const unit = String(item.unit || '').toUpperCase();
    if (unit !== targetUnit) continue;
    const qty = parseFloat(String(item.quantity || ''));
    if (isNaN(qty) || qty <= 0) continue;
    qtySum += qty;
    hasQty = true;
  }

  if (isThermo) {
    return {
      markingTypes: 'N/A',
      quantityCompleted: hasQty ? qtySum : null,
      paintMaterial: 'N/A',
    };
  }

  // MMA: collect distinct categories and materials from completed items
  const cats = new Set<string>();
  const mats = new Set<string>();
  for (const item of items) {
    if (String(item.status || '').toLowerCase() !== 'completed') continue;
    const cat = String(item.category || '').trim();
    const mat = String(item.colorMaterial || '').trim();
    if (cat) cats.add(cat);
    if (mat && mat.toLowerCase() !== 'n/a') mats.add(mat);
  }

  return {
    markingTypes: Array.from(cats).sort().join(', '),
    quantityCompleted: hasQty ? qtySum : null,
    paintMaterial: Array.from(mats).sort().join(', '),
  };
}
