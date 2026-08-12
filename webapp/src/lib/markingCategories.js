// Single source of truth for Marking Type metadata used across the
// React app. Keep in sync with the mirror CATEGORY_UNITS map in
// Apps Script Code.js (setupMarkingItems() / seedMarkingItems_ /
// handleCreateMarkingItem_ / handleUpdateMarkingItem_).

export const MARKING_CATEGORIES = [
  // WO Top Table
  'Double Yellow Line', 'Double White Line', 'Lane Lines', 'Gores', 'Messages', 'Arrows',
  'Solid Lines', 'Rail Road X/Diamond', 'Others',
  // Intersection Grid
  'HVX Crosswalk', 'Stop Msg', 'Stop Line',
  // Page 2 detailed lines
  '4" Line', '6" Line', '8" Line', '12" Line', '16" Line', '24" Line',
  // Page 2 messages
  'Only Msg', 'Bus Msg', 'Bump Msg', 'Custom Msg', '20 MPH Msg',
  // Page 2 railroad
  'Railroad (RR)', 'Railroad (X)',
  // Page 2 arrows
  'L/R Arrow', 'Straight Arrow', 'Combination Arrow', 'Combination Arrow (L/R)',
  // Page 2 misc
  'Speed Hump Markings', 'Shark Teeth 12x18', 'Shark Teeth 24x36',
  // Page 2 bike lane — the DOT is migrating the preform bike symbol from
  // the old rider style to a new bike-only style; both map to the same
  // Bike Symbol field on the CFR/PL and sum, but price differently.
  'Bike Lane Arrow', 'Old Bike Symbol (w/ rider)', 'New Bike Symbol (just bike)',
  'Bike Lane Green Bar',
  // Preform pedestrian symbol
  'Pedestrian Men',
  // MMA
  'Bike Lane', 'Pedestrian Space', 'Bus Lane',
]

// Categories rendered with the grid layout: Type | Intersection | Direction | Qty | Unit
export const GRID_CATEGORIES = new Set(['HVX Crosswalk', 'Stop Msg', 'Stop Line'])

// Categories rendered with the MMA layout: Type | Color/Material | Qty | Unit
// (and require a Color/Material value to be considered Completable).
export const MMA_CATEGORIES = new Set(['Bike Lane', 'Bus Lane', 'Pedestrian Space'])

// Fixed unit per Marking Type. Categories omitted from the map are
// treated as "variable" — the user picks SF/LF/EA manually. Only
// "Others" is intentionally variable today.
export const CATEGORY_UNITS = {
  // ── Square Feet (MMA area work) ──────────────────────────────
  'Bike Lane':           'SF',
  'Bus Lane':            'SF',
  'Pedestrian Space':    'SF',

  // ── Linear Feet (lines, crosswalks, stop lines) ──────────────
  'Double Yellow Line':  'LF',
  'Double White Line':   'LF',
  'Lane Lines':          'EA',   // 10' skips — each EA is one 10 LF skip of 4" line
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

  // ── Each / count (messages, arrows, misc) ────────────────────
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
  'Bike Lane Symbol':    'EA',   // legacy — retired from picker, kept for old rows
  'Old Bike Symbol (w/ rider)':  'EA',
  'New Bike Symbol (just bike)': 'EA',
  'Pedestrian Men':      'EA',
  'Bike Lane Green Bar': 'SF',

  // "Others" is intentionally absent — variable unit.
}

export const UNIT_OPTIONS = ['SF', 'LF', 'EA']

// Fixed Color/Material per MMA Marking Type. The color is a property of
// the marking itself — a bike lane is always green, a bus lane always
// red, a pedestrian space always truffle — so it's derived from the
// category rather than typed. Categories omitted here fall back to the
// MMA_COLORS picker (only reachable when work_type is MMA but the
// category isn't one of the three). Keep in sync with the mirror
// CATEGORY_COLORS_ map in Apps Script Code.js.
export const CATEGORY_COLORS = {
  'Bike Lane':        'Green',
  'Bus Lane':         'Red',
  'Pedestrian Space': 'Truffle',
}

export const MMA_COLORS = ['Green', 'Truffle', 'Red']

export function colorForCategory(category) {
  return CATEGORY_COLORS[category] || null
}

export function colorIsLocked(category) {
  return CATEGORY_COLORS[category] != null
}

// Display-only aliases. The KEY is the canonical stored/routing string
// (unchanged everywhere — sheet value, pricing, unit, and PDF routing all
// still key on it); the VALUE is only what the crew sees in the UI. Apply
// via displayCategory() at render leaves ONLY, never in state/payloads/keys.
// Editing a value here changes the visible label and nothing else.
export const CATEGORY_DISPLAY_OVERRIDES = {
  'Speed Hump Markings': 'Speed Bump Arrow',
  'Lane Lines': "Lane Lines (10' Skips)",
}

// Presentation transform: canonical category string -> what the user sees.
// Falls through to the raw string for everything not aliased.
export function displayCategory(category) {
  return CATEGORY_DISPLAY_OVERRIDES[category] || category
}

export function unitForCategory(category) {
  return CATEGORY_UNITS[category] || null
}

export function unitIsLocked(category) {
  return CATEGORY_UNITS[category] != null
}

export function pickLayout(item) {
  const cat = item.category || ''
  if (GRID_CATEGORIES.has(cat)) return 'grid'
  if (item.section === 'Intersection Grid') return 'grid'
  if (MMA_CATEGORIES.has(cat)) return 'mma'
  if (String(item.work_type || '').toLowerCase() === 'mma') return 'mma'
  return 'default'
}

export function rowRequiresColor(item) {
  if (MMA_CATEGORIES.has(item.category || '')) return true
  return String(item.work_type || '').toLowerCase() === 'mma'
}

// The color a row actually carries: derived from the Marking Type when
// that type has a fixed color, else whatever is stored. Legacy rows
// written before colors were derived may have a blank or free-text
// col K — those still read as Green/Red/Truffle here, which is what the
// UI shows and what the server rewrites on the next save.
export function effectiveColor(item) {
  return colorForCategory(item.category || '')
    || String(item.color_material || '').trim()
}

export function rowIsCompletable(item) {
  if (item.status === 'Completed') return true
  const qty = parseFloat(item.quantity)
  if (isNaN(qty) || qty <= 0) return false
  if (!String(item.unit || '').trim()) return false
  if (rowRequiresColor(item) && !effectiveColor(item)) return false
  return true
}
