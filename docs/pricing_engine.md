# Pricing engine — canonical reference

How revenue gets calculated for each Marking Item. Source of truth for
the tables below lives in two places that must stay in sync:

- **Apps Script** — `Code.js` "PRICING ENGINE" section
  (`PRICING_GROUP_BY_CATEGORY_`, `LINE_WIDTH_MULTIPLIER_`,
  `LINE12_MULTIPLIER_`, `EXTRUDED_UNIT_COUNT_`, `PREFORMED_UNIT_COUNT_`,
  `priceMarkingItem_`).
- **React webapp** — `webapp/src/lib/pricing.js` mirrors the same tables
  so the UI can render group labels and flag unpriced items without
  re-fetching. Revenue math itself always runs server-side.

Cross-reference with [`marking_type_mapping.md`](./marking_type_mapping.md)
for how categories roll up onto the CFR top table and Production Log.

## Pricing groups

Each Marking Items category is routed to exactly one group. The group
determines the formula and which Contract Pricing column supplies the
base rate.

| Group | Formula | Contract Pricing column |
|---|---|---|
| `line4` | `qty × line4_rate × LINE_WIDTH_MULTIPLIER[cat]` | E (line4 base $/LF) |
| `line12` | `qty × line12_rate × LINE12_MULTIPLIER[cat]` | F (line12 base $/LF) |
| `preformed` | `qty × preformed_rate × PREFORMED_UNIT_COUNT[cat]` | G ($/Unit) |
| `extruded` | `qty × extruded_rate × EXTRUDED_UNIT_COUNT[cat]` | H ($/Unit) |
| `color_surface` | `qty × color_surface_rate` | I ($/SF) |
| `unpriced` | n/a — surfaces in **Needs Pricing** bucket | n/a |

The Contract Pricing row is resolved by (contractor, contract number,
borough, date_completed). Dated rows win over blank-date rows when both
match. When no row matches, the item is flagged `no_rate`.

## `line4` — line widths × $/LF base rate

| Category | Multiplier | Unit |
|---|---|---|
| 4" Line | 1.0 | LF |
| 6" Line | 1.5 | LF |
| 8" Line | 2.0 | LF |
| 12" Line | 3.0 | LF |
| 16" Line | 4.0 | LF |
| 24" Line | 6.0 | LF |
| Lane Lines | 1.0 | LF |
| Double Yellow Line | 2.0 | LF |

Standard width / 4 across all contractors. Double Yellow gets 2.0
because it's a paired stripe; Lane Lines is treated as a 4" base.

## `line12` — crosswalk / stop line $/LF base rate

| Category | Multiplier | Unit |
|---|---|---|
| HVX Crosswalk | 1.0 | LF |
| Stop Line | 2.0 | LF |

Stop Line is a 24" stripe and bills at 2× the 12" crosswalk base rate.
HVX Crosswalk uses the base rate unchanged.

## `preformed` — preformed thermoplastic ($/Unit × units)

| Category | Units | Unit |
|---|---|---|
| Bike Lane Symbol | 0.91 | EA |

## `extruded` — standard NYC DOT thermo unit table ($/Unit × units)

Same unit count across all contractors per the DOT chart.

### Messages — sums of per-letter unit counts from the 8' Letters & Numbers table

| Category | Units | Derivation |
|---|---|---|
| Stop Msg | 1.35 | S(0.37)+T(0.25)+O(0.39)+P(0.34) |
| Only Msg | 1.35 | O(0.39)+N(0.46)+L(0.25)+Y(0.25) |
| Bus Msg | 1.19 | B(0.46)+U(0.36)+S(0.37) |
| Bump Msg | 1.64 | B(0.46)+U(0.36)+M(0.48)+P(0.34) |
| 20 MPH Msg | 1.97 | 2(0.37)+0(0.39)+M(0.48)+P(0.34)+H(0.39) |
| Railroad (RR) | 0.82 | R(0.41)+R(0.41) |
| Railroad (X) | 0.31 | X(0.31) |

### Symbols / arrows — direct from the Symbols (Extruded) DOT table

| Category | Units | DOT chart name |
|---|---|---|
| L/R Arrow | 1.00 | Turn Arrow |
| Straight Arrow | 0.81 | Through (Straight) Arrow |
| Combination Arrow | 1.65 | Combo Arrow |
| Combination Arrow (L/R) | 1.74 | Combo Arrow (left/right) |
| Bike Lane Arrow | 0.29 | Bicycle Facility Arrow |
| Speed Hump Markings | 0.78 | Speed Hump Marking |
| Shark Teeth 12x18 | 0.05 | Sharks Teeth 12" x 18" |
| Shark Teeth 24x36 | 0.19 | Sharks Teeth 24" x 36" |

**Note on Combination Arrow (L/R):** priced separately (1.74 vs 1.65)
but rolls up into the same `Combination Arrows` cell on the CFR top
table and the same `Combination Arrow` row on the Production Log —
admins don't need separate line items for the two flavors.

## `color_surface` — surface treatments ($/SF base rate)

No per-category multiplier; quantity is in SF.

| Category | Unit |
|---|---|
| Bike Lane | SF |
| Bus Lane | SF |
| Pedestrian Space | SF |
| Bike Lane Green Bar | SF *(EA entries are blocked with `unit_migration` until re-entered as SF)* |

## `unpriced` — surfaces in the Needs Pricing bucket

Items in these categories always need manual pricing review.

| Category | Why it's unpriced |
|---|---|
| Custom Msg | Variable text — unit count differs per message, can't be tabled |
| Others | Catch-all bucket — by definition no fixed pricing |
| Solid Lines | Parent — re-classify as a specific line width (`4" Line`–`24" Line`) before pricing |
| Gores | Parent — typically reported as `12" Line` |
| Messages | Parent — re-classify as the specific message type |
| Arrows | Parent — re-classify as the specific arrow type |
| Rail Road X/Diamond | Parent — re-classify as `Railroad (RR)` or `Railroad (X)` |

## Needs Pricing reason codes

When `priceMarkingItem_` can't produce a number, it returns one of these
reasons so the UI can group and explain them.

| Reason | Meaning |
|---|---|
| `unit_migration` | Bike Lane Green Bar entered as EA (legacy); re-enter as SF |
| `unpriced_category` | Category is in the always-manual bucket |
| `no_rate` | No Contract Pricing row matches (contractor, contract, borough) |
| `no_unit_count` | Extruded / preformed / line12 multiplier missing from the table |

## DOT chart items we don't price yet

These appear in the DOT Symbols (Extruded) chart but aren't in
`MARKING_CATEGORIES`, so they can't be entered as marking items today.
Add them if Oneiro starts performing this work.

| DOT chart name | Units | Group it would belong to |
|---|---|---|
| Lane Reduction Arrow | 2.71 | extruded |
| Wrong Way Arrow | 1.57 | extruded |
| HOV Lane | 0.87 | extruded |
| Bike Symbol 40" x 72" | N.A. (preformed) | preformed |
| Bike Symbol 24" x 48" | N.A. (preformed) | preformed |
| Ped Symbol 72" | N.A. (preformed) | preformed |
