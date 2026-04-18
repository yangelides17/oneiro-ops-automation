"""
parse_work_order.py — Claude Vision parser for NYC DOT Work Order scans
========================================================================
Called by watch_and_fill.py when a new PDF lands in the Drive Scan Inbox.

Flow:
  1. watch_and_fill.py downloads the PDF from Drive
  2. Calls parse(file_bytes, mime_type) here
  3. normalize_wo_data() cleans and derives missing fields
  4. Returns a dict ready to write to the Work Order Tracker sheet

Environment variables required:
  ANTHROPIC_API_KEY  — Claude API key (set in Railway dashboard)

Work Order Tracker column mapping (0-indexed):
  0   Work Order #
  1   Prime Contractor
  2   Contract Number
  3   Borough                  ← normalized from WO boro code (K→BK etc.)
  4   Contract ID / Reg #      ← blank on intake; admin fills from prime contractor
  5   Location
  6   From Street
  7   To Street
  8   Due Date
  9   Priority Level
  10  Pavement Work Type
  11  WO Received Date         ← "Issue To Contractor Date" on the WO form
  12  Water Blast Required?    ← derived: "Yes - MMA" / "NA" / "Pending"
  13  Water Blast Confirmed?   ← always "No" on intake
  14  Water Blast SQFT         ← extracted if handwritten; else blank
  15  Status                   ← "Received" on intake
  16  Dispatch Date            ← blank
  17  Work Start Date          ← blank
  18  Work End Date            ← blank
  19  Marking Types            ← formatted string from extracted marking rows
  20  SQFT Completed           ← blank (filled from field report)
  21  Paint / Material Used    ← blank (filled from field report)
  22  Issues Reported          ← blank
  23  Photos Uploaded?         ← "No"
  24  Production Log Done?     ← "No"
  25  Field Report Done?       ← "No"
  26  Invoice #                ← blank
  27  Invoice Date             ← blank
  28  Invoice Amount           ← blank
  29  Invoice Sent?            ← "No"
  30  Payment Received?        ← "No"
  31  Payment Date             ← blank
  32  Certified Payroll Week   ← blank
  33  Filed?                   ← "No"
  34  Notes                    ← general remarks from WO
"""

import base64
import json
import os
import logging

log = logging.getLogger(__name__)

# ── Borough code normalization ────────────────────────────────────────────────
# WO form uses single-letter codes; the Tracker uses the 2-letter abbreviations
BOROUGH_MAP = {
    'K':  'BK',   # Brooklyn (WO form uses K)
    'BK': 'BK',
    'M':  'M',    # Manhattan (tracker uses M, not MN)
    'MN': 'M',
    'BX': 'BX',   # Bronx
    'Q':  'QU',   # Queens
    'QU': 'QU',
    'SI': 'SI',   # Staten Island
    'R':  'SI',   # alternate code for Staten Island (Richmond)
}

# Words in General Remarks that indicate MMA work (requires waterblasting)
MMA_REMARK_KEYWORDS = [
    'bike lane',
    'bus lane',
    'pedestrian space',
    'pedestrian stop',
    'ped space',
    'ped stop',
    'mma',
    'color surface',
]

# ── Extraction prompt ─────────────────────────────────────────────────────────
EXTRACTION_PROMPT = """You are parsing a scanned NYC DOT Pavement Marking Work Order Report form.
The form is issued by NYC DOT and given to pavement marking contractors like Metro Express.

Extract every field listed below and return ONLY valid JSON — no explanation, no markdown fences.
If a field is not visible, illegible, or not applicable, use null. Never guess.

{
  "work_order_id":      "The Work Order number, e.g. PT-11930. Usually in top-left area labeled 'Work Order'.",
  "contractor":         "Prime contractor name, e.g. METRO EXPRESS. Labeled 'Contractor'.",
  "contract_number":    "Full contract number including suffix, e.g. 84122MBTP496/SFT. Labeled 'Contract Number'.",
  "borough":            "Single-letter or abbreviated boro code exactly as printed: K, M, BX, Q or QU, SI, R. Labeled 'Boro'.",
  "location":           "Street name in the 'Location' field.",
  "from_street":        "Street name in the 'From/At' field.",
  "to_street":          "Street name in the 'To' field.",
  "due_date":           "Due Date as printed on the form, e.g. 8/28/2025.",
  "priority_level":     "Priority level label, e.g. '3 - Schedule'. Labeled 'Priority Level'.",
  "pavement_work_type": "Pavement work type, e.g. REFURBISHMENT or NEW. Labeled 'Pavement Work'.",
  "wo_received_date":   "The 'Issue To Contractor Date' at the bottom of the form, e.g. 07/17/2025.",
  "water_blast_sqft":   "If any waterblasting square footage is handwritten anywhere on the form, extract the number as an integer. Otherwise null.",
  "general_remarks":    "The full text of the General Remarks section (middle of the form, labeled 'General Remarks>>>>>'). Transcribe exactly as written, including handwritten text.",

  "top_markings": [
    {
      "category":    "One of: Double Yellow Line, Lane Lines, Gores, Messages, Arrows, Solid Lines, Rail Road X/Diamond, Others. Use the label as printed in the middle of the top table (a row per line/marking type).",
      "description": "The free-text description to the right of the category label, e.g. 'RECAP FROM HAMILTON PL TO 2ND AV'. Transcribe exactly, including 'RECAP' if present. Leave out any rows where the description column is blank."
    }
  ],

  "intersection_grid": [
    {
      "intersection": "The intersection name as printed in the leftmost INTERSECTIONS column, e.g. '5 AV', 'HAMILTON PL'. Transcribe exactly.",
      "n":          "The value of the 'North' column for this row, usually 'HVX' or blank.",
      "e":          "The value of the 'East' column for this row.",
      "s":          "The value of the 'South' column for this row.",
      "w":          "The value of the 'West' column for this row.",
      "stop_msg":   "The value of the 'Stop Msg' column for this row. Usually blank, or a directional string like 'West', 'East', 'EW', 'NSEW'.",
      "stop_lines": "The value of the 'Stop lines' (far-right) column for this row. Same format as stop_msg."
    }
  ]
}

Important notes:
- water_blast_sqft may appear as a handwritten number near the words 'waterblast', 'WB', or 'water blast' anywhere on the form.
- The General Remarks section is critical — it often contains handwritten notes about the type of work (e.g. 'RECAP PAINT FOR BIKE LANE', 'BUS LANE', 'PED SPACE'). Transcribe it fully.
- The Issue To Contractor Date is near the bottom of the form.

top_markings rules:
- This is the upper table that lists marking CATEGORIES down the middle column (Double Yellow CenterLine / Lane Lines / Gores / Messages / Arrows / Solid Lines / Rail Road X / Diamond / Others).
- Only include a row if its description column contains any text (e.g. 'RECAP', 'RECAP FROM HAMILTON PL TO 2ND AV'). Skip blank rows entirely.
- Normalize the category label: output 'Double Yellow Line' (not 'Double Yellow CenterLine'), 'Rail Road X/Diamond' (not 'Rail Road X / Diamond').
- Preserve the description text verbatim including all caps.
- Order top_markings by their printed row order.

intersection_grid rules:
- This is the bottom table with column headers: INTERSECTIONS | Order | North | East | South | West | Stop Msg | Sch M 8' | Sch M 10' | Stop lines
- IGNORE the Order, Sch M 8', and Sch M 10' columns — they are unused in our system.
- Only include a row if the INTERSECTIONS cell has text AND at least one of N/E/S/W/stop_msg/stop_lines has a non-empty value. Skip blank rows at the bottom of the table.
- For N/E/S/W cells: copy the value verbatim. Typical value is 'HVX' when a crosswalk is required at that direction; blank otherwise.
- For stop_msg and stop_lines cells: copy the directional string verbatim. Values are usually 'North', 'East', 'South', 'West', or concatenations like 'EW' (East AND West), 'NS', 'NSEW'. Preserve the exact letters.
- Return an empty array if the form has no intersection grid entries.
- Order intersection_grid top-to-bottom as printed.
"""


def parse(file_bytes: bytes, mime_type: str = 'application/pdf') -> dict:
    """
    Send a scanned WO file to Claude Vision and return normalized tracker row data.

    Args:
        file_bytes: raw bytes of the PDF or image file
        mime_type:  'application/pdf', 'image/jpeg', 'image/png', etc.

    Returns:
        dict with all WO Tracker fields populated, or {'_parse_error': reason} on failure.
    """
    try:
        import anthropic
    except ImportError:
        raise RuntimeError("anthropic package not installed — run: pip install anthropic")

    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY environment variable not set")

    client = anthropic.Anthropic(api_key=api_key)

    encoded = base64.standard_b64encode(file_bytes).decode('utf-8')

    if mime_type == 'application/pdf':
        content_block = {
            "type": "document",
            "source": {"type": "base64", "media_type": mime_type, "data": encoded}
        }
    else:
        content_block = {
            "type": "image",
            "source": {"type": "base64", "media_type": mime_type, "data": encoded}
        }

    try:
        response = client.messages.create(
            model="claude-opus-4-6",
            # Bumped from 1024 to 4096 — Thermo WOs can have 10+ intersection
            # grid rows, each with 7 fields, plus the top-markings array.
            max_tokens=4096,
            messages=[{
                "role": "user",
                "content": [
                    content_block,
                    {"type": "text", "text": EXTRACTION_PROMPT}
                ]
            }]
        )

        raw = response.content[0].text.strip()

        # Strip markdown code fences if present
        if raw.startswith('```'):
            lines = raw.split('\n')
            raw = '\n'.join(lines[1:-1] if lines[-1] == '```' else lines[1:])

        extracted = json.loads(raw)
        return normalize_wo_data(extracted)

    except json.JSONDecodeError as e:
        log.error(f"Claude returned non-JSON: {e}")
        return {'_parse_error': f'JSON decode failed: {e}'}
    except Exception as e:
        log.error(f"Claude API call failed: {e}")
        return {'_parse_error': str(e)}


def normalize_wo_data(raw: dict) -> dict:
    """
    Clean and derive all fields needed for the Work Order Tracker row.

    - Normalizes borough code (K→BK, M→MN, etc.)
    - Formats marking_types as a readable string
    - Derives water_blast_required from work type / marking types
    - Builds the full 35-column row ready for appendRow()
    """
    # ── Borough normalization ─────────────────────────────────────
    raw_borough = str(raw.get('borough') or '').strip().upper()
    borough = BOROUGH_MAP.get(raw_borough, raw_borough)

    # ── Marking types → readable string ──────────────────────────
    marking_list = raw.get('marking_types') or []
    marking_parts = []
    for m in marking_list:
        if not isinstance(m, dict):
            continue
        mtype = (m.get('type') or '').strip()
        mval  = (str(m.get('value') or '')).strip()
        if mtype:
            marking_parts.append(f"{mtype}: {mval}" if mval else mtype)
    marking_types_str = ', '.join(marking_parts)

    # ── Water blast logic ─────────────────────────────────────────
    # Valid dropdown values in the tracker:
    #   Water Blast Required?  → "Yes - MMA" | "No - Thermo" | "N/A" | "" (blank if unsure)
    #   Water Blast Confirmed? → "Yes" | "No" | "N/A"
    #
    # Detection priority:
    #   1. Explicit WB SQFT handwritten on the form → "Yes - MMA"
    #   2. General remarks mention Bike Lane / Bus Lane / Ped Space/Stop → "Yes - MMA"
    #   3. Remarks present but no MMA indicators → "No - Thermo"
    #   4. No remarks at all → "" (blank — admin decides)

    wb_sqft_raw = raw.get('water_blast_sqft')
    wb_sqft = ''
    if wb_sqft_raw is not None:
        try:
            wb_sqft = int(float(str(wb_sqft_raw).replace(',', '')))
        except (ValueError, TypeError):
            wb_sqft = str(wb_sqft_raw).strip()

    remarks = (raw.get('general_remarks') or '').strip()
    remarks_lower = remarks.lower()
    is_mma = bool(wb_sqft) or any(kw in remarks_lower for kw in MMA_REMARK_KEYWORDS)
    has_remarks = bool(remarks)

    if is_mma:
        water_blast_required = 'Yes - MMA'
        # Confirmed only if we saw explicit WB info on the form
        water_blast_confirmed = 'Yes' if wb_sqft else 'No'
    elif has_remarks:
        # Remarks present but no MMA indicators → thermoplastic work
        water_blast_required  = 'No - Thermo'
        water_blast_confirmed = 'N/A'
        wb_sqft = ''  # not applicable
    else:
        # No remarks to go on — leave blank for admin to decide
        water_blast_required  = ''
        water_blast_confirmed = 'N/A'
        wb_sqft = ''

    # ── Normalize top_markings + intersection_grid ────────────────
    # Both are passed through to the Apps Script handler, which expands them
    # into individual Marking Items rows (one row per discrete piece of work).
    top_markings = []
    for m in (raw.get('top_markings') or []):
        if not isinstance(m, dict):
            continue
        category = (m.get('category') or '').strip()
        description = (m.get('description') or '').strip()
        if category and description:
            top_markings.append({'category': category, 'description': description})

    intersection_grid = []
    for ig in (raw.get('intersection_grid') or []):
        if not isinstance(ig, dict):
            continue
        intersection = (ig.get('intersection') or '').strip()
        if not intersection:
            continue
        cells = {
            'n':          (ig.get('n') or '').strip(),
            'e':          (ig.get('e') or '').strip(),
            's':          (ig.get('s') or '').strip(),
            'w':          (ig.get('w') or '').strip(),
            'stop_msg':   (ig.get('stop_msg') or '').strip(),
            'stop_lines': (ig.get('stop_lines') or '').strip(),
        }
        # Only include if at least one direction/stop cell is populated.
        if any(v for v in cells.values()):
            intersection_grid.append({'intersection': intersection, **cells})

    # ── Derive work_type (MMA vs Thermo) ──────────────────────────
    # MMA detection above already sets water_blast_required = 'Yes - MMA'
    # when remarks/handwritten WB SQFT indicate MMA work. Intersection
    # grids and top-marking "RECAP" entries are Thermo-specific, so if
    # either is populated we default to Thermo. Admin can override later.
    if water_blast_required == 'Yes - MMA':
        work_type = 'MMA'
    elif intersection_grid or top_markings:
        work_type = 'Thermo'
    else:
        work_type = ''   # admin decides

    # ── Build normalized dict ─────────────────────────────────────
    return {
        'work_order_id':         (raw.get('work_order_id') or '').strip(),
        'prime_contractor':      (raw.get('contractor') or '').strip().title(),
        'contract_number':       (raw.get('contract_number') or '').strip(),
        'borough':               borough,
        'location':              (raw.get('location') or '').strip().title(),
        'from_street':           (raw.get('from_street') or '').strip().title(),
        'to_street':             (raw.get('to_street') or '').strip().title(),
        'due_date':              (raw.get('due_date') or '').strip(),
        'priority_level':        (raw.get('priority_level') or '').strip(),
        'pavement_work_type':    (raw.get('pavement_work_type') or '').strip().upper(),
        'wo_received_date':      (raw.get('wo_received_date') or '').strip(),
        'water_blast_required':  water_blast_required,
        'water_blast_confirmed': water_blast_confirmed,
        'water_blast_sqft':      wb_sqft,
        'status':                'Received',
        'work_type':             work_type,
        'top_markings':          top_markings,
        'intersection_grid':     intersection_grid,
    }
