"""
fill_contractor_field_report.py
-------------------------------
Fills the NYC DOT Contractor Field Report PDF template for a completed
Work Order. Called by watch_and_fill.py when a JSON payload with
`_type = 'contractor_field_report'` lands in the Field Reports Drive
folder.

Template: Thermo Contractor Field Report Template_FORM.pdf
Triggered by: handleSubmitFieldReport_ in Code.js when wo_complete=true.

Expected data shape (any key may be missing — blank fields stay blank):
{
  "_type": "contractor_field_report",
  "wo_id":            "RM-43281",
  "date_entered":     "4/13/2026",       # from WO scan (parse_work_order)
  "work_order":       "RM-43281",
  "contractor":       "Metro Express",
  "contract_number":  "84125MBTP701/EXT",
  "boro":             "BK",
  "location":         "13 ST",
  "school":           "NA",               # from WO scan; default "NA"
  "from":             "5 AV",
  "to":               "Hamilton Pl",
  "install_from":     "4/18/2026",
  "install_to":       "4/20/2026",
  "general_remarks":  "...",              # aggregated Issues Reported across all submits
  "markings": {                           # top pavement-markings table
    "Double Yellow Line": 260,
    "Lane Line":          180,
    ...
  },
  "grid": [                               # intersection grid, 10 rows max
    {"intersection": "5 AV", "order": "", "n": "", "e": "", "s": "",
     "w": 60, "stop_msg": "", "sch8": "", "sch10": "", "st_line": ""},
    ...
  ],
  "prep_by":          "Mankaryous, Beshoy"
}

Fields intentionally NOT filled (out of scope — blank on output PDF):
  Work Type, Associated WO, Traffic #, Punch dates, Order # / Drawing # per
  marking row, Road Condition radio, Crew Chief, Contractor Notes,
  Liquidated Damages, entire bottom signature section (prime contractor
  signs physically).

Implementation note — template quirk:
  The CFR template's widget annotations (on the page) and terminal fields
  (under /AcroForm/Fields) are SEPARATE PDF objects (not merged, as in
  the Production Log template). pypdf's update_page_form_field_values()
  writes /V on the widget, but viewers read text-field values from the
  terminal field. So we bypass the helper and walk /AcroForm/Fields
  directly. See _write_fields_via_acroform() below.
"""

import json
import logging
import os
import sys
import warnings

logging.getLogger('pypdf').setLevel(logging.ERROR)
warnings.filterwarnings('ignore', message='.*Font dictionary.*not found.*', module='pypdf')

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    NameObject, BooleanObject, DictionaryObject, IndirectObject, create_string_object
)

TEMPLATE = os.path.join(
    os.path.dirname(__file__), '..', 'templates',
    'Thermo_Contractor_Field_Report_Fillable.pdf'
)

# ── Top pavement-markings table: category → value field name ──────
# Order and Drawing columns intentionally omitted (out of scope).
# Category names use the SAME strings the UI/Marking Items sheet uses,
# so the Apps Script aggregator can emit them directly. Known renames
# from Marking Items are handled here:
#   Marking Items "Lane Lines" (plural)      → "Lane Line"  (singular)
#   Marking Items "Combination Arrow"         → "Combination Arrows"
#   Marking Items "L/R Arrow"                 → "L/R Arrows"
#   Marking Items "Straight Arrow"            → "Straight Arrows"
MARKING_TABLE = {
    'Double Yellow Line':  'Text_1',
    'Lane Line':           'Text_2',
    '4" Line':             'Text_3',
    '6" Line':             'page0_field31',
    '8" Line':             'Text_4',
    '12" Line':            'Text_5',
    '16" Line':            'Text_6',
    '24" Line':            'page0_field41',
    'Stop Msg':            'page0_field23',
    'Only Msg':            'page0_field26',
    'Bus Msg':             'Text_7',
    'Bump Msg':            'page0_field32',
    'Custom Msg':          'Text_8',
    '20 MPH Msg':          'Text_9',
    'Railroad (RR)':       'Text_10',
    'Railroad (X)':        'Text_11',
    'L/R Arrows':          'Text_12',
    'Straight Arrows':     'Text_13',
    'Combination Arrows':  'Text_14',
    'Speed Hump Markings': 'Text_15',
    'Shark Teeth 12x18':   'Text_16',
    'Shark Teeth 24x36':   'Text_17',
    'Bike Lane Arrows':    'Text_18',
    'Bike Lane Symbols':   'Text_19',
    'Bike Lane Green Bar': 'Text_20',
}

# ── Intersection grid rows (10 columns each) ───────────────────────
# Column order: intersection | order | N | E | S | W | stop_msg | sch8 | sch10 | st_line
GRID_ROWS = [
    ('page0_field46','page0_field47','page0_field48','page0_field49','page0_field50','page0_field51','page0_field52','page0_field53','page0_field54','page0_field55'),
    ('page0_field56','page0_field57','page0_field58','page0_field59','page0_field60','page0_field61','page0_field62','page0_field63','page0_field64','page0_field65'),
    ('page0_field66','page0_field67','page0_field68','page0_field69','page0_field70','page0_field71','page0_field72','page0_field73','page0_field74','page0_field75'),
    ('page0_field76','page0_field77','page0_field78','page0_field79','page0_field80','page0_field81','page0_field82','page0_field83','page0_field84','page0_field85'),
    ('page0_field86','page0_field87','page0_field88','page0_field89','page0_field90','page0_field91','page0_field92','page0_field93','page0_field94','page0_field95'),
    ('page0_field96','page0_field97','page0_field98','page0_field99','page0_field100','page0_field101','page0_field102','page0_field103','page0_field104','page0_field105'),
    ('page0_field106','page0_field107','page0_field108','page0_field109','page0_field110','page0_field111','page0_field112','page0_field113','page0_field114','page0_field115'),
    ('page0_field116','page0_field117','page0_field118','page0_field119','page0_field120','page0_field121','page0_field122','page0_field123','page0_field124','page0_field125'),
    ('page0_field126','page0_field127','page0_field128','page0_field129','page0_field130','page0_field131','page0_field132','page0_field133','page0_field134','page0_field135'),
    ('page0_field136','page0_field137','page0_field138','page0_field139','page0_field140','page0_field141','page0_field142','page0_field143','page0_field144','page0_field145'),
]
GRID_COL_KEYS = ('intersection', 'order', 'n', 'e', 's', 'w', 'stop_msg', 'sch8', 'sch10', 'st_line')


def build_field_map(data: dict) -> dict:
    """Flatten the structured data dict into {field_name: value_str}."""
    f = {}

    # Header
    f['page0_field1']  = data.get('date_entered', '')
    f['page0_field2']  = data.get('work_order', '')
    f['page0_field3']  = data.get('contractor', '')
    f['page0_field4']  = data.get('contract_number', '')
    # page0_field5 (Work Type) — out of scope, leave blank
    f['page0_field6']  = data.get('boro', '')
    # page0_field7 (Associated WO), page0_field8 (Traffic #) — out of scope
    f['page0_field9']  = data.get('location', '')
    f['page0_field10'] = data.get('school', '')
    f['page0_field11'] = data.get('from', '')
    f['page0_field12'] = data.get('to', '')

    # Installation dates
    f['page0_field13'] = data.get('install_from', '')
    f['page0_field14'] = data.get('install_to', '')
    # page0_field15/16 (Punch dates) — out of scope

    # Pavement-markings table
    markings = data.get('markings') or {}
    for category, field_name in MARKING_TABLE.items():
        v = markings.get(category)
        if v not in (None, ''):
            f[field_name] = str(v)

    # General Remarks (new dedicated field added to template)
    f['Text_21'] = data.get('general_remarks', '')

    # Intersection grid
    for row_fields, row_data in zip(GRID_ROWS, data.get('grid') or []):
        if not isinstance(row_data, dict):
            continue
        for key, field_name in zip(GRID_COL_KEYS, row_fields):
            v = row_data.get(key)
            if v not in (None, ''):
                f[field_name] = str(v)

    # Bottom block — only Prep By in scope
    f['page0_field146'] = data.get('prep_by', '')

    # Filter out empties so we don't overwrite pre-existing template text
    return {k: v for k, v in f.items() if v not in (None, '')}


def _write_fields_via_acroform(writer, payload):
    """
    This CFR template stores widget annotations (on the page) and
    terminal AcroForm fields as SEPARATE objects with matching /T names
    — they're not merged via /Parent like in most templates. Different
    viewers read /V from different places:

      * Adobe / Preview read /V from the terminal field and regenerate
        appearances when /AcroForm /NeedAppearances is true.
      * pdf.js (and therefore react-pdf in our webapp) reads /V from
        the widget annotation and does NOT regenerate appearances —
        it renders whatever /V + /AP the widget carries.

    Write /V on BOTH so every viewer shows the filled values.
    """
    acroform = writer._root_object.get('/AcroForm')
    if not acroform:
        return
    fields = acroform.get('/Fields')
    if not fields:
        return

    # Terminal fields keyed by /T
    field_by_name = {}
    for fref in fields:
        f = fref.get_object() if isinstance(fref, IndirectObject) else fref
        t = f.get('/T')
        if t:
            field_by_name[str(t)] = f

    # Widget annotations keyed by /T across every page. Built once so
    # we don't re-scan pages per-field.
    widgets_by_name = {}
    for page in writer.pages:
        annots = page.get('/Annots')
        if annots is None:
            continue
        if isinstance(annots, IndirectObject):
            annots = annots.get_object()
        for aref in annots:
            a = aref.get_object() if isinstance(aref, IndirectObject) else aref
            if not isinstance(a, DictionaryObject):
                continue
            if a.get('/Subtype') != '/Widget':
                continue
            t = a.get('/T')
            if not t:
                continue
            widgets_by_name.setdefault(str(t), []).append(a)

    for name, value in payload.items():
        v = create_string_object(str(value))

        # Terminal field (Adobe/Preview read here)
        f = field_by_name.get(name)
        if f is not None and f.get('/FT') == '/Tx':
            f[NameObject('/V')] = v

        # Widget annotations with the same /T (pdf.js reads here)
        for widget in widgets_by_name.get(name, []):
            widget[NameObject('/V')] = v
            # If an existing appearance stream exists, delete it so the
            # viewer falls back to rendering /V. pdf.js otherwise keeps
            # showing the stale blank appearance.
            if '/AP' in widget:
                del widget[NameObject('/AP')]


def fill(data: dict, template_path: str = TEMPLATE, output_path: str = None) -> str:
    """Fill the template with data and write to output_path. Returns the path."""
    if output_path is None:
        wo  = data.get('work_order', 'unknown')
        dt  = str(data.get('install_to') or data.get('date_entered') or 'unknown').replace('/', '-')
        output_path = f'Contractor_Field_Report_{wo}_{dt}_FILLED.pdf'

    reader = PdfReader(template_path)
    writer = PdfWriter(clone_from=reader)

    payload = build_field_map(data)
    _write_fields_via_acroform(writer, payload)

    if '/AcroForm' in writer._root_object:
        writer._root_object['/AcroForm'].update({
            NameObject('/NeedAppearances'): BooleanObject(True)
        })

    with open(output_path, 'wb') as f:
        writer.write(f)
    return output_path


if __name__ == '__main__':
    if len(sys.argv) > 1:
        with open(sys.argv[1]) as fh:
            data = json.load(fh)
        out = fill(data, output_path=sys.argv[2] if len(sys.argv) > 2 else None)
        print(f'\u2713 Filled PDF written \u2192 {out}')
    else:
        # Quick smoke test with in-scope dummy data
        sample = {
            'wo_id':            'RM-43281',
            'date_entered':     '4/13/2026',
            'work_order':       'RM-43281',
            'contractor':       'Metro Express',
            'contract_number':  '84125MBTP701/EXT',
            'boro':             'BK',
            'location':         '13 ST',
            'school':           'NA',
            'from':             '5 AV',
            'to':                'Hamilton Pl',
            'install_from':     '4/18/2026',
            'install_to':       '4/20/2026',
            'general_remarks':  'Recap from Hamilton Pl to 2 AV on Double Yellow.',
            'markings': {
                'Double Yellow Line': 260,
                'Lane Line':          180,
                'Stop Msg':           4,
                'L/R Arrows':         6,
            },
            'grid': [
                {'intersection': '5 AV',        'w': 60},
                {'intersection': '4 AV',        'e': 55, 'w': 55, 'st_line': 40},
                {'intersection': '3 AV',        'e': 60, 'w': 60, 'st_line': 40},
                {'intersection': '2 AV',        'e': 65, 'w': 65, 'stop_msg': 2, 'st_line': 50},
                {'intersection': 'HAMILTON PL', 'e': 70, 'stop_msg': 1, 'st_line': 25},
            ],
            'prep_by':          'Mankaryous, Beshoy',
        }
        out = fill(sample, output_path='/tmp/cfr_smoke.pdf')
        print(f'\u2713 Filled PDF written \u2192 {out}')
