// Single source of truth for pricing-group lookups used by the
// Revenue tab. Keep in sync with Code.js — see PRICING_GROUP_BY_CATEGORY_,
// LINE_WIDTH_MULTIPLIER_, EXTRUDED_UNIT_COUNT_, PREFORMED_UNIT_COUNT_
// in the Apps Script "PRICING ENGINE" section.
//
// Revenue math runs SERVER-SIDE only; this module exists so the React
// side can render group labels, badge unmapped categories, and (later)
// preview line-item invoice values without re-fetching the server.

export const PRICING_GROUPS = ['line4', 'line12', 'preformed', 'extruded', 'color_surface']

export const PRICING_GROUP_BY_CATEGORY = {
  // line4 — base $/LF × LINE_WIDTH_MULTIPLIER
  '4" Line':            'line4',
  '6" Line':            'line4',
  '8" Line':            'line4',
  '12" Line':           'line4',
  '16" Line':           'line4',
  '24" Line':           'line4',
  'Lane Lines':         'line4',
  'Double Yellow Line': 'line4',

  // line12 — Crosswalk + Stop Line bulk rate
  'HVX Crosswalk':      'line12',
  'Stop Line':          'line12',

  // preformed thermoplastic
  'Bike Lane Symbol':   'preformed',

  // extruded thermoplastic ($/Unit × EXTRUDED_UNIT_COUNT)
  'Stop Msg':            'extruded',
  'Only Msg':            'extruded',
  'Bus Msg':             'extruded',
  'Bump Msg':            'extruded',
  '20 MPH Msg':          'extruded',
  'Railroad (RR)':       'extruded',
  'Railroad (X)':        'extruded',
  'L/R Arrow':           'extruded',
  'Straight Arrow':      'extruded',
  'Combination Arrow':   'extruded',
  'Speed Hump Markings': 'extruded',
  'Shark Teeth 12x18':   'extruded',
  'Shark Teeth 24x36':   'extruded',
  'Bike Lane Arrow':     'extruded',

  // color surface treatment ($/SF)
  'Bike Lane':           'color_surface',
  'Bus Lane':            'color_surface',
  'Pedestrian Space':    'color_surface',
  'Bike Lane Green Bar': 'color_surface',

  // always unpriced — surface in Needs Pricing bucket
  'Custom Msg':          'unpriced',
  'Others':              'unpriced',
  'Solid Lines':         'unpriced',
  'Gores':               'unpriced',
  'Messages':            'unpriced',
  'Arrows':              'unpriced',
  'Rail Road X/Diamond': 'unpriced',
}

// Width / 4 — standard across contractors per the user's pricing convention.
export const LINE_WIDTH_MULTIPLIER = {
  '4" Line':            1.0,
  '6" Line':            1.5,
  '8" Line':            2.0,
  '12" Line':           3.0,
  '16" Line':           4.0,
  '24" Line':           6.0,
  'Lane Lines':         1.0,
  'Double Yellow Line': 2.0,
}

// Standard NYC DOT thermo unit table. `null` = unit count not yet known —
// item lands in the Needs Pricing bucket with reason='no_unit_count'
// until the contractor unit table arrives. Mirror of EXTRUDED_UNIT_COUNT_
// in Code.js — see that block for the per-letter derivations.
export const EXTRUDED_UNIT_COUNT = {
  'Stop Msg':            1.35,
  'Only Msg':            1.35,
  'Bus Msg':             1.19,
  'Bump Msg':            1.64,
  '20 MPH Msg':          1.97,
  'Railroad (RR)':       0.82,
  'Railroad (X)':        0.31,
  'L/R Arrow':           1.00,
  'Straight Arrow':      0.81,
  'Combination Arrow':   1.65,
  'Speed Hump Markings': 0.78,
  'Shark Teeth 12x18':   0.05,
  'Shark Teeth 24x36':   0.19,
  'Bike Lane Arrow':     0.29,
}

export const PREFORMED_UNIT_COUNT = {
  'Bike Lane Symbol': 1.0,
}

export const PRICING_GROUP_LABEL = {
  line4:         '4" Line group',
  line12:        'Crosswalk / Stop Line',
  preformed:     'Preformed L&S',
  extruded:      'Extruded L&S',
  color_surface: 'Color Surface',
  unpriced:      'Unpriced',
}

export const NEEDS_PRICING_REASON_LABEL = {
  no_rate:           'No Contract Pricing row matches this contract',
  no_unit_count:     'Unit count missing from extruded table',
  unpriced_category: 'Category requires manual pricing',
  unit_migration:    'Bike Lane Green Bar entered as EA — re-enter as SF',
}
