# Marking Type mapping — canonical reference

Living cross-reference of every **Marking Items** category vs the
**Contractor Field Report (CFR)** top table vs the **Production Log
(PL)** marking grid. Use this when reviewing discrepancies with the
ops manager or when we add new categories. Items marked `⚠️` still
need ops-manager confirmation.

CFR is the source of truth for category names; the PL template's
printed labels drift slightly and get reconciled via a rename map
inside `Code.js`.

## Confirmed mappings

| Marking Items category | Unit | CFR top-table label | PL row label                     |
|------------------------|------|---------------------|----------------------------------|
| Double Yellow Line     | LF   | Double Yellow Line  | Double Yellow Line (Center Line) |
| Lane Lines             | LF   | Lane Line           | Lane Lines 4" (Skips)            |
| 4" Line                | LF   | 4" Line             | 4" Lines                         |
| 6" Line                | LF   | 6" Line             | 6" Lines                         |
| 8" Line                | LF   | 8" Line             | 8" Lines                         |
| 12" Line               | LF   | 12" Line            | 12" Lines (Gore)                 |
| 16" Line               | LF   | 16" Line            | 16" Lines                        |
| 24" Line               | LF   | 24" Line            | 24" Lines                        |
| HVX Crosswalk          | LF   | (goes into grid)    | CrossWalks/Stop Lines *(sum with Stop Line)* |
| Stop Line              | LF   | (goes into grid)    | CrossWalks/Stop Lines *(sum with HVX Crosswalk)* |
| Stop Msg               | EA   | Stop Msg            | Stop Message                     |
| Only Msg               | EA   | Only Msg            | Message Only                     |
| Bus Msg                | EA   | Bus Msg             | Bus Message                      |
| Bump Msg               | EA   | Bump Msg            | Bump                             |
| 20 MPH Msg             | EA   | 20 MPH Msg          | 20 MPH Message                   |
| Railroad (RR)          | EA   | Railroad (RR)       | Railroad - RR                    |
| Railroad (X)           | EA   | Railroad (X)        | Railroad - X                     |
| L/R Arrow              | EA   | L/R Arrows          | Left & or Right Arrows           |
| Straight Arrow         | EA   | Straight Arrows     | Straight Arrow                   |
| Combination Arrow      | EA   | Combination Arrows  | Combination Arrow                |
| Speed Hump Markings    | EA   | Speed Hump Markings | Speed Hump Marking               |
| Shark Teeth 24x36      | EA   | Shark Teeth 24x36   | Sharks Teeth 24" 36"             |
| Bike Lane Arrow        | EA   | Bike Lane Arrows    | Bicycle Lane Arrow               |
| Bike Lane Symbol       | EA   | Bike Lane Symbols   | Bicycle Lane Symbol              |

## MMA SF categories — NOT in the grid; flow through Color Surface Treatment rows

Aggregation rule:
- **Color Surface Treatment 1** = sum of all Marking Items Quantity
  where Unit = `SF` for the WO (Bike Lane + Bus Lane + Pedestrian
  Space, usually just one per WO).
- **Color Surface Treatment 2** = distinct `Color/Material` values
  from those rows, comma-joined.

| Marking Items category | Unit | CFR top-table label | PL destination                         |
|------------------------|------|---------------------|----------------------------------------|
| Bike Lane              | SF   | (none)              | Color Surface Treatment 1 (SQFT sum) + 2 (color) |
| Bus Lane               | SF   | (none)              | Color Surface Treatment 1 (SQFT sum) + 2 (color) |
| Pedestrian Space       | SF   | (none)              | Color Surface Treatment 1 (SQFT sum) + 2 (color) |

## Intentionally blank on the PL

- **Custom Msg** (Marking Items, EA) — on CFR as `Custom Msg`. No PL
  row; CFR is authoritative.
- **Water Blasting for Surface** (PL row 25) — Oneiro doesn't do
  waterblasting currently. Admin can hand-fill if it ever changes.

## Parent-only categories (never carry quantities)

Top table of the original WO uses these as *type* markers that tell
the crew "some marking of this type is needed"; reporting always
happens at the specific subtype level. These categories therefore
never receive quantity entries:

- `Gores` → reported as specific widths (usually `12" Line`)
- `Messages` → reported as `Stop Msg` / `Only Msg` / `Bus Msg` / etc.
- `Arrows` → reported as `L/R Arrow` / `Straight Arrow` / `Combination Arrow`
- `Solid Lines` → reported as specific widths (`4" Line` – `24" Line`)
- `Rail Road X/Diamond` → reported as `Railroad (X)` / `Railroad (RR)`

`Double Yellow Line` and `Lane Lines` are exceptions — top-level
reporting categories with no subtypes.

Whether to hide these parent categories from the Marking Items
dropdown so crews can't accidentally pick them is an ops-manager
call for a later pass.

## Tentatively left off the PL — ops manager to confirm ⚠️

Implementation leaves these blank on the PL. If the ops manager
wants them captured there, we add PL template rows in a follow-up
and extend `PL_CATEGORY_MAP_`.

- **Shark Teeth 12x18** (Marking Items, EA) — on CFR. No PL row.
- **Bike Lane Green Bar** (Marking Items, EA) — on CFR. No PL row.

## Still open for ops-manager review ⚠️

1. **`Ped Stop`** (Marking Items, EA) — NOT on the CFR. Origin is
   unclear. The PL has a `PED X-ING Message` row (pedestrian crossing
   message). Is `Ped Stop` the same thing as `PED X-ING Message`, or
   are they different? Both stay unmapped until ops confirms.
2. **`PED X-ING Message`** (PL row 24) — no obvious Marking Items
   source. Paired with the `Ped Stop` question above. Leave blank
   for now.
