"""
fill_production_log.py
----------------------
Fills Metro Thermoplastic Production Log PDF template from a JSON data file
(produced by Apps Script) or from supplied data dict directly.

Usage:
  python3 fill_production_log.py <data.json>   # from Apps Script export
  OR import fill_production_log and call fill(data, template_path, output_path)

Field mapping (Production Log):
  Header:
    page0_field1  = Crew #
    page0_field2  = Truck #
    Date          = Date (MM/dd/yyyy)
    page0_field3  = Inspector Present (Y/N)
    page0_field4  = Refilled Gas Tank (Y/N)

  Crew (crew chief + up to 4 individuals):
    crew_chief_name, in_chief, out_chief
    Text_2 = Thermo (White) bags used
    Text_3 = Thermo (Yellow) bags used
    Text_4 = Beads (bags)
    Text_5 = Paint (cans)
    Text_6 = (spare material field)
    crew_name_{1-4}, in_{1-4}, out_{1-4}

  Per Work Order (columns 1–4):
    WO_number_{n}               = Work order #
    Location_{n}                = Location/street
    borough_bk_{n}              = "O" if Brooklyn, else ""
    borough_qu_{n}              = "O" if Queens, else ""
    borough_bx_{n}              = "O" if Bronx, else ""
    borough_m_{n}               = "O" if Manhattan, else ""
    color_surface_treat_line_1_{n} = Square footage with "SQFT" suffix (e.g. "450 SQFT")
    color_surface_treat_line_2_{n} = Paint color
    WO_complete_{n}             = "Y" or "N"
    page0_field140–143          = Layout Y/N per WO column
    page0_field144–147          = Layout hours per WO column

  Marking grid (page0_field26–126, 25 rows × 4 cols):
    row_labels = [
      'Double Yellow Line (Center Line)', 'Lane Lines 4" (Skips)',
      '4" Lines', '6" Lines', '8" Lines', '12" Lines (Gore)',
      '16" Lines', '24" Lines', 'Railroad - X', 'CrossWalks/Stop Lines',
      'Speed Hump Marking', 'Sharks Teeth 24" 36"', 'Bus Message',
      'Message Only', 'Stop Message', 'Left & or Right Arrows',
      'Straight Arrow', 'Combination Arrow', 'Railroad - RR', 'Bump',
      'Bicycle Lane Arrow', 'Bicycle Lane Symbol', '20 MPH Message',
      'PCO #/VNG Message', 'Water Blasting for Surface'
    ]
    grid_fields[row][col] = page0_field{26 + row*4 + col}
    (col order matches WO column order: 1, 2, 3, 4)
    NOTE: field38 is skipped (gap in original numbering)
"""

import json
import re
import sys
import os
import logging
import warnings

# pypdf emits a non-fatal warning when a PDF references ArialMT (a common
# Windows font) that isn't in its internal metrics table; it falls back to
# Helvetica automatically.  The filled output is correct — suppress the noise.
logging.getLogger('pypdf').setLevel(logging.ERROR)
warnings.filterwarnings('ignore', message='.*Font dictionary.*not found.*', module='pypdf')

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    NameObject, ArrayObject, FloatObject,
    DecodedStreamObject, DictionaryObject, create_string_object
)

TEMPLATE = os.path.join(
    os.path.dirname(__file__),
    "templates/Metro_Production_Log_Fillable.pdf"
)

# Borough code → field suffix mapping
BOROUGH_MAP = {
    'BK': 'bk', 'BROOKLYN': 'bk',
    'QU': 'qu', 'QUEENS': 'qu',
    'BX': 'bx', 'BRONX': 'bx',
    'MN': 'm',  'MANHATTAN': 'm', 'M': 'm',
    'SI': '',   'STATEN ISLAND': '',   # no circle field for SI
}

# Marking grid row labels (25 rows)
MARKING_ROWS = [
    'Double Yellow Line (Center Line)', 'Lane Lines 4" (Skips)',
    '4" Lines', '6" Lines', '8" Lines', '12" Lines (Gore)',
    '16" Lines', '24" Lines', 'Railroad - X', 'CrossWalks/Stop Lines',
    'Speed Hump Marking', 'Sharks Teeth 24" 36"', 'Bus Message',
    'Message Only', 'Stop Message', 'Left & or Right Arrows',
    'Straight Arrow', 'Combination Arrow', 'Railroad - RR', 'Bump',
    'Bicycle Lane Arrow', 'Bicycle Lane Symbol', '20 MPH Message',
    'PCO #/VNG Message', 'Water Blasting for Surface'
]

def _grid_field_name(row_idx, col_idx):
    """Return the page0_fieldX name for marking grid position (0-indexed row, col)."""
    n = 26 + row_idx * 4 + col_idx
    if n >= 38:   # gap: field38 is missing from original template
        n += 1
    return f'page0_field{n}'


def build_field_map(data: dict) -> dict:
    """
    Convert a structured data dict into a flat {field_name: value} dict
    ready for update_page_form_field_values().

    Expected data shape (all keys optional — unset fields stay blank):
    {
      "date":             "04/10/2026",
      "crew_number":      "1",
      "truck_number":     "2",
      "inspector_present": "Y",
      "gas_tank_refilled": "Y",
      "materials": {
          "thermo_white_bags":  "12",
          "thermo_yellow_bags": "0",
          "beads_bags":         "4",
          "paint_cans":         "0"
      },
      "crew_chief": { "name": "Carlos Rivera", "time_in": "7:00 AM", "time_out": "3:30 PM" },
      "crew": [
          { "name": "Jose Martinez", "time_in": "7:00 AM", "time_out": "3:30 PM" },
          ...
      ],
      "work_orders": [
          {
            "wo_number":  "PT-1234",
            "borough":    "QU",
            "location":   "QUEENS BLVD & 82 ST",
            "sqft":       "450",
            "paint":      "White",
            "complete":   "Y",
            "layout_yn":  "N",
            "layout_hours": "",
            "markings": {
                "4\" Lines": "250",
                "CrossWalks/Stop Lines": "200"
            }
          },
          ...   (max 4 work orders)
      ]
    }
    """
    f = {}

    # ── Header ────────────────────────────────────────────────────────────────
    f['Date']          = str(data.get('date', ''))
    f['page0_field1']  = str(data.get('crew_number', ''))
    f['page0_field2']  = str(data.get('truck_number', ''))
    f['page0_field3']  = str(data.get('inspector_present', ''))
    f['page0_field4']  = str(data.get('gas_tank_refilled', ''))

    # ── Materials ─────────────────────────────────────────────────────────────
    mats = data.get('materials', {})
    f['Text_2'] = str(mats.get('thermo_white_bags', ''))
    f['Text_3'] = str(mats.get('thermo_yellow_bags', ''))
    f['Text_4'] = str(mats.get('beads_bags', ''))
    f['Text_5'] = str(mats.get('paint_cans', ''))
    f['Text_6'] = str(mats.get('other', ''))

    # ── Crew chief ────────────────────────────────────────────────────────────
    chief = data.get('crew_chief', {})
    f['crew_chief_name'] = str(chief.get('name', ''))
    f['in_chief']        = str(chief.get('time_in', ''))
    f['out_chief']       = str(chief.get('time_out', ''))

    # ── Individual crew members (max 4) ───────────────────────────────────────
    crew = data.get('crew', [])
    for i, member in enumerate(crew[:4], start=1):
        f[f'crew_name_{i}'] = str(member.get('name', ''))
        f[f'in_{i}']        = str(member.get('time_in', ''))
        f[f'out_{i}']       = str(member.get('time_out', ''))

    # ── Work orders (max 4 columns) ───────────────────────────────────────────
    wos = data.get('work_orders', [])
    for n, wo in enumerate(wos[:4], start=1):
        f[f'WO_number_{n}'] = str(wo.get('wo_number', ''))
        f[f'Location_{n}']  = str(wo.get('location', ''))
        f[f'WO_complete_{n}'] = str(wo.get('complete', ''))

        # Color Surface Treatment (named fields)
        sqft_val = wo.get('sqft', '')
        f[f'color_surface_treat_line_1_{n}'] = f"{sqft_val} SQFT" if sqft_val else ''
        f[f'color_surface_treat_line_2_{n}'] = str(wo.get('paint', ''))

        # Borough circles  — write "O" in the correct field, blank in others
        borough_key = str(wo.get('borough', '')).upper()
        borough_suffix = BOROUGH_MAP.get(borough_key, '')
        for suffix in ['bk', 'qu', 'bx', 'm']:
            val = 'O' if suffix == borough_suffix else ''
            f[f'borough_{suffix}_{n}'] = val

        # Layout rows (page0_field140–147: 140-143 = Y/N, 144-147 = hours)
        layout_yn_field    = f'page0_field{140 + (n-1)}'
        layout_hours_field = f'page0_field{144 + (n-1)}'
        f[layout_yn_field]    = str(wo.get('layout_yn', ''))
        f[layout_hours_field] = str(wo.get('layout_hours', ''))

        # Marking grid (25 rows)
        markings = wo.get('markings', {})
        col_idx = n - 1
        for row_idx, label in enumerate(MARKING_ROWS):
            field_name = _grid_field_name(row_idx, col_idx)
            f[field_name] = str(markings.get(label, ''))

    return f


def _patch_borough_ap(writer: PdfWriter, field_map: dict) -> None:
    """
    The Production Log template stores borough circle fields with /AP /BBox [0,0,0,0].
    pypdf's auto_regenerate sees a zero BBox and produces a stub 'q Q' appearance,
    causing the PDF viewer to fall back to unclipped /V + /DA rendering (oversized O).

    This helper manually injects a correct appearance stream for every borough field
    that has a non-empty value.  The 'O' is drawn at BOROUGH_FONT_PT, centered inside
    the field's actual rect.

    Helvetica capital-O metrics (approximate):
      advance width  ≈ 0.722 × pt
      cap height     ≈ 0.728 × pt
      descender      ≈ 0.0 pt  (O has no descender)
    """
    BOROUGH_FONT_PT = 30          # tweak here if size needs further adjustment

    # Glyph metrics for Helvetica 'O'
    ADV_W   = 0.722 * BOROUGH_FONT_PT
    CAP_H   = 0.728 * BOROUGH_FONT_PT

    from pypdf.generic import IndirectObject

    def resolve(obj):
        while isinstance(obj, IndirectObject):
            obj = writer._reader.get_object(obj) if hasattr(writer, '_reader') else obj.get_object()
        return obj

    for page in writer.pages:
        raw = page.get('/Annots')
        if not raw:
            continue
        for ref in raw:
            if isinstance(ref, IndirectObject):
                annot = ref.get_object()
            else:
                annot = ref
            if not isinstance(annot, DictionaryObject):
                continue

            t = str(annot.get('/T', ''))
            if not t.startswith('borough_'):
                continue

            # Only patch fields we're writing a value into
            val = field_map.get(t, '')
            if not val:
                # Ensure /V is cleared and AP shows blank
                annot[NameObject('/V')] = create_string_object('')
                continue

            # Compute field dimensions from /Rect
            rect = annot.get('/Rect', ArrayObject([FloatObject(0)]*4))
            x0, y0, x1, y1 = [float(v) for v in rect]
            fw = x1 - x0   # field width  in pts
            fh = y1 - y0   # field height in pts

            # Center the glyph
            tx = (fw - ADV_W) / 2
            ty = (fh - CAP_H) / 2

            # Build the content stream
            content = (
                f"q\n"
                f"/Tx BMC\n"
                f"q\n"
                f"0 0 {fw:.3f} {fh:.3f} re\n"
                f"W\n"
                f"BT\n"
                f"/Helvetica {BOROUGH_FONT_PT} Tf 0 g\n"
                f"{tx:.3f} {ty:.3f} Td\n"
                f"(O) Tj\n"
                f"ET\n"
                f"Q\n"
                f"EMC\n"
                f"Q\n"
            ).encode()

            # Build appearance XObject with font resource so the viewer can render text
            font_obj = DictionaryObject({
                NameObject('/Type'):     NameObject('/Font'),
                NameObject('/Subtype'):  NameObject('/Type1'),
                NameObject('/BaseFont'): NameObject('/Helvetica'),
            })
            resources = DictionaryObject({
                NameObject('/Font'): DictionaryObject({
                    NameObject('/Helvetica'): font_obj
                })
            })

            ap_stream = DecodedStreamObject()
            ap_stream.set_data(content)
            ap_stream[NameObject('/Type')]      = NameObject('/XObject')
            ap_stream[NameObject('/Subtype')]   = NameObject('/Form')
            ap_stream[NameObject('/BBox')]      = ArrayObject([
                FloatObject(0), FloatObject(0),
                FloatObject(fw), FloatObject(fh)
            ])
            ap_stream[NameObject('/Resources')] = resources

            # Stream objects MUST be indirect per PDF spec — add to writer object table
            ap_ref = writer._add_object(ap_stream)

            # Wire into annotation /AP /N using the indirect reference
            ap_dict = DictionaryObject()
            ap_dict[NameObject('/N')] = ap_ref
            annot[NameObject('/AP')] = ap_dict
            annot[NameObject('/V')]  = create_string_object(val)
            # Also update /DA to match our explicit size
            annot[NameObject('/DA')] = create_string_object(
                f'/Helvetica {BOROUGH_FONT_PT} Tf 0 g'
            )


def fill(data: dict, template_path: str = TEMPLATE, output_path: str = None) -> str:
    """Fill template with data dict and write to output_path. Returns output path."""
    if output_path is None:
        date_str = data.get('date', 'unknown').replace('/', '-')
        output_path = f'Production_Log_{date_str}_FILLED.pdf'

    reader = PdfReader(template_path)
    writer = PdfWriter()
    writer.append(reader)

    field_map = build_field_map(data)

    # 1. Fill all non-borough fields via pypdf's standard path
    non_borough = {k: v for k, v in field_map.items() if not k.startswith('borough_')}
    writer.update_page_form_field_values(writer.pages[0], non_borough)

    # 2. Borough fields get manually-constructed appearances (correct BBox + centered O)
    _patch_borough_ap(writer, field_map)

    with open(output_path, 'wb') as fh:
        writer.write(fh)
    return output_path


# ── CLI entry point ───────────────────────────────────────────────────────────
if __name__ == '__main__':
    if len(sys.argv) > 1:
        with open(sys.argv[1]) as fh:
            data = json.load(fh)
        out = fill(data, output_path=sys.argv[2] if len(sys.argv) > 2 else None)
        print(f'✅  Filled PDF written → {out}')
    else:
        # ── Sample / test data ────────────────────────────────────────────────
        sample = {
            "date": "04/10/2026",
            "crew_number": "1",
            "truck_number": "T-07",
            "inspector_present": "Y",
            "gas_tank_refilled": "Y",
            "materials": {
                "thermo_white_bags": "14",
                "thermo_yellow_bags": "2",
                "beads_bags": "6",
                "paint_cans": "0"
            },
            "crew_chief": {
                "name": "Carlos Rivera",
                "time_in": "7:00 AM",
                "time_out": "3:30 PM"
            },
            "crew": [
                {"name": "Jose Martinez",  "time_in": "7:00 AM", "time_out": "3:30 PM"},
                {"name": "Maria Lopez",    "time_in": "7:00 AM", "time_out": "3:30 PM"},
                {"name": "David Chen",     "time_in": "7:30 AM", "time_out": "3:30 PM"},
            ],
            "work_orders": [
                {
                    "wo_number":    "PT-1234",
                    "borough":      "QU",
                    "location":     "QUEENS BLVD & 82 ST",
                    "sqft":         "450",
                    "paint":        "White",
                    "complete":     "Y",
                    "layout_yn":    "N",
                    "layout_hours": "",
                    "markings": {
                        '4" Lines': "180",
                        'CrossWalks/Stop Lines': "270"
                    }
                },
                {
                    "wo_number":    "PT-1235",
                    "borough":      "BK",
                    "location":     "FLATBUSH AVE & CHURCH AVE",
                    "sqft":         "320",
                    "paint":        "Yellow",
                    "complete":     "N",
                    "layout_yn":    "Y",
                    "layout_hours": "1.5",
                    "markings": {
                        'Double Yellow Line (Center Line)': "320"
                    }
                },
            ]
        }

        out = fill(sample, output_path='test_production_log_filled.pdf')
        print(f'✅  Test fill written → {out}')
