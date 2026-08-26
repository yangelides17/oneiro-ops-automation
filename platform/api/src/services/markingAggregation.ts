/**
 * Marking item aggregation for Production Log generation.
 *
 * Ported from Code.js aggregateMarkingItemsForPL_ (lines 927–1021),
 * PL_CATEGORY_MAP_ (lines 864–900), PL_FOLD_MULTIPLIER_ (lines 908–910).
 *
 * Converts raw marking items into the shape the Production Log PDF filler
 * expects: a dict of PL row labels → quantities, plus MMA SF rollup.
 */

/**
 * Maps marking item categories to Production Log row labels.
 * Categories not in this map are handled specially (HVX, Stop Line, Ped Men, MMA SF)
 * or dropped silently.
 */
export const PL_CATEGORY_MAP: Record<string, string> = {
  // LF
  'Double Yellow Line':  'Double Yellow Line (Center Line)',
  'Double White Line':   '4" Lines',   // folds at 2× qty (see PL_FOLD_MULTIPLIER)
  'Lane Lines':          'Lane Lines 4" (Skips)',
  '4" Line':             '4" Lines',
  '6" Line':             '6" Lines',
  '8" Line':             '8" Lines',
  '12" Line':            '12" Lines (Gore)',
  '16" Line':            '16" Lines',
  '24" Line':            '24" Lines',
  // EA
  'Stop Msg':            'Stop Message',
  'Only Msg':            'Message Only',
  'Bus Msg':             'Bus Message',
  'Bump Msg':            'Bump',
  '20 MPH Msg':          '20 MPH Message',
  'Railroad (RR)':       'Railroad - RR',
  'Railroad (X)':        'Railroad - X',
  'L/R Arrow':           'Left & or Right Arrows',
  'Straight Arrow':      'Straight Arrow',
  'Combination Arrow':   'Combination Arrow',
  'Combination Arrow (L/R)': 'Combination Arrow', // rolls into same row
  'Speed Hump Markings': 'Speed Hump Marking',
  'Shark Teeth 24x36':   'Sharks Teeth 24" 36"',
  'Bike Lane Arrow':     'Bicycle Lane Arrow',
  'Bike Lane Symbol':    'Bicycle Lane Symbol',  // legacy alias
  'Old Bike Symbol (w/ rider)':  'Bicycle Lane Symbol',
  'New Bike Symbol (just bike)': 'Bicycle Lane Symbol',
};

/**
 * Categories whose quantity is multiplied when folded into a shared PL row.
 * Double White Line prints at 2× into the 4" Lines row.
 */
export const PL_FOLD_MULTIPLIER: Record<string, number> = {
  'Double White Line': 2,
};

/** MMA surface categories that roll into the Color Surface Treatment row. */
const MMA_SF_CATEGORIES = new Set([
  'Bike Lane', 'Bus Lane', 'Pedestrian Space', 'Bike Lane Green Bar',
]);

export interface MarkingItemRow {
  category: string;
  quantity: number | string | null;
  unit: string | null;
  colorMaterial: string | null;
  dateCompleted: string | null;
  status: string;
  crewChief: string | null;
}

export interface PlAggregation {
  /** PL row label → quantity (number) or formatted string (for CrossWalks/Stop Lines). */
  markings: Record<string, number | string>;
  /** Total MMA square footage, or empty string. */
  sqft: number | string;
  /** Comma-joined colors/materials for Color Surface Treatment line 2. */
  paint: string;
}

/**
 * Aggregate completed marking items for a single WO into Production Log shape.
 * Exact port of Code.js aggregateMarkingItemsForPL_.
 *
 * @param items - All marking items for the WO
 * @param targetDate - Filter to items completed on this date (YYYY-MM-DD), or null for all
 * @param crewChief - Filter to items tagged with this crew chief, or empty for untagged items
 */
export function aggregateForProductionLog(
  items: MarkingItemRow[],
  targetDate: string | null,
  crewChief: string | null,
): PlAggregation {
  const out: PlAggregation = { markings: {}, sqft: '', paint: '' };
  const tgt = String(targetDate || '').slice(0, 10);
  const chiefFilter = String(crewChief || '').trim();

  // Filter to completed items matching date and crew chief
  const filtered = items.filter(r => {
    if (String(r.status || '').toLowerCase() !== 'completed') return false;
    if (tgt && String(r.dateCompleted || '').slice(0, 10) !== tgt) return false;
    if (String(r.crewChief || '').trim() !== chiefFilter) return false;
    return true;
  });

  if (filtered.length === 0) return out;

  let sfSum = 0;
  const colorsSet: Record<string, boolean> = {};
  let crosswalkSum = 0;
  let stopLineSum = 0;
  let pedMenSum = 0;

  for (const r of filtered) {
    const category = String(r.category || '').trim();
    const qty = Number(r.quantity);
    if (isNaN(qty) || qty <= 0) continue;
    const unit = String(r.unit || '').toUpperCase();

    // MMA SF → Color Surface Treatment, not the grid
    if (unit === 'SF' && MMA_SF_CATEGORIES.has(category)) {
      sfSum += qty;
      const color = String(r.colorMaterial || '').trim();
      if (color && color.toLowerCase() !== 'n/a') colorsSet[color] = true;
      continue;
    }

    // HVX Crosswalk + Stop Line share a row, tracked separately
    if (category === 'HVX Crosswalk') { crosswalkSum += qty; continue; }
    if (category === 'Stop Line') { stopLineSum += qty; continue; }

    // Pedestrian Men → PED X-ING Message row
    if (category === 'Pedestrian Men') { pedMenSum += qty; continue; }

    // Standard rename via PL_CATEGORY_MAP
    const plLabel = PL_CATEGORY_MAP[category];
    if (!plLabel) continue; // unmapped categories drop silently

    const foldMult = PL_FOLD_MULTIPLIER[category] || 1;
    out.markings[plLabel] = ((out.markings[plLabel] as number) || 0) + qty * foldMult;
  }

  // Compose the special rows
  if (crosswalkSum > 0 || stopLineSum > 0) {
    out.markings['CrossWalks/Stop Lines'] = `${crosswalkSum}/${stopLineSum}`;
  }
  if (pedMenSum > 0) {
    out.markings['PED X-ING Message'] = `${pedMenSum} PED MEN`;
  }
  if (sfSum > 0) {
    out.sqft = sfSum;
    out.paint = Object.keys(colorsSet).sort().join(', ');
  }

  return out;
}
