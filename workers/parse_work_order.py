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
    'K':  'BK',   # Brooklyn
    'BK': 'BK',
    'M':  'MN',   # Manhattan
    'MN': 'MN',
    'BX': 'BX',   # Bronx
    'Q':  'QU',   # Queens
    'QU': 'QU',
    'SI': 'SI',   # Staten Island
    'R':  'SI',   # alternate code for Staten Island (Richmond)
}

# Marking types that indicate MMA (methyl methacrylate) work — requires waterblast prep
MMA_KEYWORDS = [
    'color surface treatment',
    'color surface',
    'mma',
    'methyl methacrylate',
    'cst',
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
  "marking_types": [
    {
      "type": "Marking row label, e.g. 'Double Yellow CenterLine', 'Lane Lines', 'Gores', 'Color Surface Treatment', 'Arrows', etc.",
      "value": "Any handwritten quantity, SQFT, or note written next to that row. null if blank."
    }
  ],
  "water_blast_sqft":   "If any waterblasting square footage is handwritten anywhere on the form, extract the number. Otherwise null.",
  "general_remarks":    "Any text in the General Remarks section.",
  "notes":              "Any other relevant handwritten or printed notes not captured above."
}

Important notes:
- The marking_types array should include ONLY rows that have something written next to them (a number, date, or note).
- Do not include blank marking rows.
- water_blast_sqft may appear as a handwritten number anywhere on the form near the word 'waterblast', 'WB', or 'water blast'.
- The Issue To Contractor Date may be labeled 'Issue To Contractor Date' near the bottom of the form.
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
            max_tokens=1024,
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
    # Waterblasting is only required for MMA work; it is N/A for thermo.
    # If waterblast SQFT is handwritten on the WO, that confirms it's needed.
    wb_sqft_raw = raw.get('water_blast_sqft')
    wb_sqft = ''
    if wb_sqft_raw is not None:
        try:
            wb_sqft = int(float(str(wb_sqft_raw).replace(',', '')))
        except (ValueError, TypeError):
            wb_sqft = str(wb_sqft_raw).strip()

    # Detect MMA from marking types
    marking_text = marking_types_str.lower()
    is_mma = any(kw in marking_text for kw in MMA_KEYWORDS)

    if wb_sqft:
        # Explicit waterblast quantity found on form
        water_blast_required = 'Yes - MMA'
    elif is_mma:
        # MMA work detected but no explicit WB notation — needs confirmation
        water_blast_required = 'Pending'
    else:
        # Thermo or non-MMA work — waterblast not applicable
        water_blast_required = 'NA'

    # ── Notes ─────────────────────────────────────────────────────
    remarks = raw.get('general_remarks') or ''
    extra   = raw.get('notes') or ''
    notes   = '; '.join(filter(None, [remarks, extra]))

    # ── Build normalized dict (matches Tracker column order) ─────
    return {
        # Fields populated from scan
        'work_order_id':      (raw.get('work_order_id') or '').strip(),
        'prime_contractor':   (raw.get('contractor') or '').strip().title(),
        'contract_number':    (raw.get('contract_number') or '').strip(),
        'borough':            borough,
        'contract_id':        '',   # blank — admin fills from prime contractor
        'location':           (raw.get('location') or '').strip().title(),
        'from_street':        (raw.get('from_street') or '').strip().title(),
        'to_street':          (raw.get('to_street') or '').strip().title(),
        'due_date':           (raw.get('due_date') or '').strip(),
        'priority_level':     (raw.get('priority_level') or '').strip(),
        'pavement_work_type': (raw.get('pavement_work_type') or '').strip().upper(),
        'wo_received_date':   (raw.get('wo_received_date') or '').strip(),
        'marking_types':      marking_types_str,
        'water_blast_required':  water_blast_required,
        'water_blast_confirmed': 'No',
        'water_blast_sqft':      wb_sqft,
        'notes':              notes,

        # Status fields — all set to intake defaults
        'status':              'Received',
    }
