"""
fill_certified_payroll.py
--------------------------
Fills NYC Certified Payroll PDF template from a structured data dict.

CORRECT field mapping (verified from AcroForm position scan):

  Header (separate fields):
    PAYROLL                                        = Payroll #
    EMPLOYER NAME                                  = Employer name
    EMPLOYER ADDRESS                               = Employer address
    EMPLOYER EMAIL ADDRESS                         = Employer email
    EMPLOYER PHONE                                 = Employer phone
    EMPLOYER TAX ID                                = Employer tax ID
    NAME OF PRIME CONTRACTOR BUILDING OWNER OR UTILITY = Prime contractor
    CONTRACT REGISTRATION                          = Contract registration #
    AGENCY                                         = Agency
    AGENCY PIN                                     = Agency PIN
    PROJECT NAME                                   = Project name
    PROJECT OR BUILDING ADDRESS                    = Project address
    Week Ending                                    = Week ending date
    PLA                                            = PLA checkbox ('/Yes' or '/Off')

  Day / Date columns (7 days, index 0–6):
    Day.{d}   = Day label ("S","M","T","W","R","F","S")  — Sun–Sat
    Date.{d}  = Date string — use "M/DD" format (no leading zero on month)

  Per-worker (N = 1–7):
    WORKER NAME ADDRESS LAST FOUR DIGITS OF SSNRow{N}  = Name / address / SSN4
    TradeClassification{N}                              = Trade classification text
    workergroup{N}                                      = J/A radio: 'Journeyperson' or 'Apprentice'

    ST hours by day:
      Worker 1:  S.0.0 (day 0)  +  S.1, S.2, S.3, S.4, S.5, S.6 (days 1–6)
      Worker 2:  S.0.2.0 – S.0.2.6
      Worker 3:  S.0.4.0 – S.0.4.6
      Worker 4:  S.0.6.0 – S.0.6.6
      Worker 5:  S.0.8.0 – S.0.8.6
      Worker 6:  S.0.10.0 – S.0.10.6
      Worker 7:  S.0.12.0.0 – S.0.12.0.6  (extra nesting in template)

    OT hours by day:
      Worker 1:  S.0.1.0 – S.0.1.6
      Worker 2:  S.0.3.0 – S.0.3.6
      Worker 3:  S.0.5.0 – S.0.5.6
      Worker 4:  S.0.7.0 – S.0.7.6
      Worker 5:  S.0.9.0 – S.0.9.6
      Worker 6:  S.0.11.0 – S.0.11.6
      Worker 7:  S.0.13.0 – S.0.13.6

    Total hours.0.{2*(N-1)}   = Worker N total ST hours
    Total hours.0.{2*(N-1)+1} = Worker N total OT hours
    Hourly Rate of Pay.{2*(N-1)}   = Worker N ST rate
    Hourly Rate of Pay.{2*(N-1)+1} = Worker N OT rate

    GROSS PAY THIS PROJECTRow{N}        = Gross pay
    Net Pay.{N-1}                       = Net pay
    Withholding and Deductions.{N-1}    = Deductions
    ANNUALIZED HOURLY RATERow{N}        = Annualized rate

  Signature:
    OFFICER OR PRINCIPAL (print)  = Signatory name
    TITLE                         = Title
    DATE                          = Date
    YEAR                          = Year
"""

import json
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
from pypdf.generic import NameObject, DictionaryObject, IndirectObject, create_string_object

TEMPLATE = os.path.join(
    os.path.dirname(__file__),
    "templates/Certified_Payroll_Fillable.pdf"
)

DEFAULT_DAY_LABELS = ["S", "M", "T", "W", "R", "F", "S"]  # Sun–Sat


def _fmt_date(date_str: str) -> str:
    """
    Format a date string for the narrow Day/Date cells.
    Strips leading zero from month only: '04/06' → '4/06', '11/06' → '11/06'
    """
    if not date_str:
        return ''
    parts = date_str.split('/')
    if len(parts) == 2:
        month = parts[0].lstrip('0') or '0'
        return f"{month}/{parts[1]}"
    return date_str


def _st_field(worker_idx: int, day: int) -> str:
    """Return the ST hours field name for worker_idx (0-based) and day (0-6)."""
    if worker_idx == 0:
        # Worker 1 is oddly split: day 0 = S.0.0, days 1–6 = S.1–S.6
        return 'S.0.0' if day == 0 else f'S.{day}'
    elif worker_idx == 6:
        # Worker 7 has an extra nesting level
        return f'S.0.12.0.{day}'
    else:
        row = 2 * worker_idx   # rows: 2,4,6,8,10 for workers 2–6
        return f'S.0.{row}.{day}'


def _ot_field(worker_idx: int, day: int) -> str:
    """Return the OT hours field name for worker_idx (0-based) and day (0-6)."""
    row = 2 * worker_idx + 1   # rows: 1,3,5,7,9,11,13 for workers 1–7
    if worker_idx == 6:
        row = 13
    return f'S.0.{row}.{day}'


def build_field_map(data: dict) -> dict:
    """
    Convert structured data dict → flat {field_name: value} for pypdf.

    Expected data shape (all keys optional):
    {
      "payroll_number":       "1",
      "week_ending":          "04/11/2026",
      "employer": {
          "name":    "Metro Thermoplastic Corp",
          "address": "123 Industrial Ave, Brooklyn NY 11201",
          "email":   "info@metrothermo.com",
          "phone":   "718-555-0100",
          "tax_id":  "12-3456789"
      },
      "prime_contractor":      "NYC Department of Transportation",
      "contract_registration": "CR-20260001",
      "agency":                "NYC DOT",
      "agency_pin":            "840MTXXXXX",
      "project_address":       "Various Streets, Queens & Brooklyn NY",
      "project_name":          "FY2026 Thermoplastic Pavement Markings",
      "pla":                   false,
      "days": [
          {"label": "M", "date": "04/06"}, ... (7 entries)
      ],
      "workers": [
          {
            "name":        "Carlos Rivera",
            "address":     "456 Oak Ave, Brooklyn NY 11203",
            "ssn4":        "1234",
            "trade":       "Pavement Marking Operator",
            "journeyperson": true,       // true=J, false=A
            "st_hours":    ["8","8","8","8","8","0","0"],
            "ot_hours":    ["0","0","0","0","2","0","0"],
            "total_st":    "40",
            "total_ot":    "2",
            "rate_st":     "35.00",
            "rate_ot":     "52.50",
            "gross_pay":   "1505.00",
            "net_pay":     "1100.00",
            "deductions":  "405.00",
            "annualized_rate": "72800.00"
          }, ...
      ],
      "signatory": {
          "name":  "Anthe Angelides",
          "title": "Principal",
          "date":  "04/14/2026",
          "year":  "2026"
      }
    }
    """
    f = {}

    # ── Header ────────────────────────────────────────────────────────────────
    f['PAYROLL']    = str(data.get('payroll_number', ''))
    f['Week Ending'] = str(data.get('week_ending', ''))

    emp = data.get('employer', {})
    f['EMPLOYER NAME']          = str(emp.get('name', ''))
    f['EMPLOYER ADDRESS']       = str(emp.get('address', ''))
    f['EMPLOYER EMAIL ADDRESS'] = str(emp.get('email', ''))
    f['EMPLOYER PHONE']         = str(emp.get('phone', ''))
    f['EMPLOYER TAX ID']        = str(emp.get('tax_id', ''))

    f['NAME OF PRIME CONTRACTOR BUILDING OWNER OR UTILITY'] = str(data.get('prime_contractor', ''))
    f['CONTRACT REGISTRATION']  = str(data.get('contract_registration', ''))
    f['AGENCY']                 = str(data.get('agency', ''))
    f['AGENCY PIN']             = str(data.get('agency_pin', ''))
    f['PROJECT OR BUILDING ADDRESS'] = str(data.get('project_address', ''))
    f['PROJECT NAME']           = str(data.get('project_name', ''))
    f['PLA']                    = '/Yes' if data.get('pla') else '/Off'

    # ── Day / Date columns ────────────────────────────────────────────────────
    days = data.get('days', [])
    for d in range(7):
        label = days[d].get('label', DEFAULT_DAY_LABELS[d]) if d < len(days) else DEFAULT_DAY_LABELS[d]
        date  = days[d].get('date', '')                      if d < len(days) else ''
        f[f'Day.{d}']  = str(label)
        f[f'Date.{d}'] = _fmt_date(date)

    # ── Workers (max 7) ───────────────────────────────────────────────────────
    workers = data.get('workers', [])
    for n, worker in enumerate(workers[:7], start=1):
        wi = n - 1   # 0-based worker index
        st_row = 2 * wi
        ot_row = 2 * wi + 1

        # Name / address / SSN4
        name_val = worker.get('name', '')
        addr_val = worker.get('address', '')
        ssn4_raw = str(worker.get('ssn4', '')).strip()
        ssn4_val = f"*** ** {ssn4_raw}" if ssn4_raw else ''
        parts = [p for p in [name_val, addr_val, ssn4_val] if p]
        f[f'WORKER NAME ADDRESS LAST FOUR DIGITS OF SSNRow{n}'] = '\n'.join(parts)

        f[f'TradeClassification{n}'] = str(worker.get('trade', ''))

        # J/A radio button
        ja = 'Journeyperson' if worker.get('journeyperson', True) else 'Apprentice'
        f[f'workergroup{n}'] = ja

        # ST hours by day
        st_hours = worker.get('st_hours', [])
        for d in range(7):
            f[_st_field(wi, d)] = str(st_hours[d]) if d < len(st_hours) else ''

        # OT hours by day
        ot_hours = worker.get('ot_hours', [])
        for d in range(7):
            f[_ot_field(wi, d)] = str(ot_hours[d]) if d < len(ot_hours) else ''

        # Totals
        f[f'Total hours.0.{st_row}'] = str(worker.get('total_st', ''))
        f[f'Total hours.0.{ot_row}'] = str(worker.get('total_ot', ''))

        # Rates
        f[f'Hourly Rate of Pay.{st_row}'] = str(worker.get('rate_st', ''))
        f[f'Hourly Rate of Pay.{ot_row}'] = str(worker.get('rate_ot', ''))

        # Pay
        f[f'GROSS PAY THIS PROJECTRow{n}']        = str(worker.get('gross_pay', ''))
        f[f'Total Gross Pay.0.0.{wi}']             = str(worker.get('total_gross_pay', ''))
        f[f'Net Pay.{wi}']                         = str(worker.get('net_pay', ''))
        f[f'Withholding and Deductions.{wi}']      = str(worker.get('deductions', ''))
        f[f'ANNUALIZED HOURLY RATERow{n}']         = str(worker.get('annualized_rate', ''))

    # ── Signature ─────────────────────────────────────────────────────────────
    sig = data.get('signatory', {})
    f['OFFICER OR PRINCIPAL (print)'] = str(sig.get('name', ''))
    f['TITLE'] = str(sig.get('title', ''))
    f['DATE']  = str(sig.get('date', ''))
    f['YEAR']  = str(sig.get('year', ''))

    return f


def _set_font_size(writer: PdfWriter, field_names: set, pt: float):
    """Set a fixed font size on specific fields BEFORE fill so pypdf uses it
    when generating appearance streams.  Also widens the /Rect by 4pt to give
    the smaller text a bit more breathing room."""
    import re
    from pypdf.generic import ArrayObject, FloatObject
    for page in writer.pages:
        for field in page.get('/Annots', []):
            field_obj = field.get_object()
            name = field_obj.get('/T')
            if name is None:
                continue
            if str(name) not in field_names:
                continue
            # Patch /DA font size
            da = str(field_obj.get('/DA', ''))
            new_da = re.sub(r'\d+(\.\d+)?\s+Tf', f'{pt} Tf', da)
            field_obj[NameObject('/DA')] = create_string_object(new_da)
            # Widen /Rect by 4pt (2pt each side) so text has room
            rect = field_obj.get('/Rect')
            if rect:
                r = [float(rect[i]) for i in range(4)]
                r[0] -= 2   # x1 left
                r[2] += 2   # x2 right
                field_obj[NameObject('/Rect')] = ArrayObject(
                    [FloatObject(v) for v in r]
                )
            # Clear cached appearance — pypdf will regenerate with new /DA
            if '/AP' in field_obj:
                del field_obj['/AP']


def _sync_checkbox_appearance_states(writer):
    """
    update_page_form_field_values sets /V on checkbox + radio-group
    fields (workergroup{N}, PLA, etc) but does NOT set /AS on the
    individual widget annotations. Adobe + Preview derive the right
    /AS from /V via /NeedAppearances; pdf.js (react-pdf in the webapp
    viewer) and Drive's preview render /AS literally and show every
    box unchecked when /AS stays at /Off.

    Walk every widget. For button-type fields (/FT=/Btn), flip /AS
    so it matches the parent field's /V:
      - If parent /V matches one of the widget's /AP/N "on" keys
        (e.g. /Journeyperson or /Apprentice), set /AS to that key.
      - Otherwise, set /AS to /Off.

    Resolves the parent via /AcroForm/Fields by /T (rather than via
    the widget's /Parent reference) so we always read the canonical
    post-write /V.

    Idempotent — safe to call after update_page_form_field_values.
    """
    def _name_str(v):
        if v is None: return None
        s = str(v)
        return s[1:] if s.startswith('/') else s

    # Build a name → terminal-field map from /AcroForm/Fields.
    terminals_by_name = {}
    acro = writer._root_object.get('/AcroForm')
    if acro is not None:
        fields = acro.get('/Fields') or []
        for fref in fields:
            f = fref.get_object() if isinstance(fref, IndirectObject) else fref
            t = f.get('/T')
            if t:
                terminals_by_name[str(t)] = f

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

            # Resolve parent name: widget's /T takes precedence, else parent's /T
            t = a.get('/T')
            if t is None:
                parent = a.get('/Parent')
                if parent is not None:
                    parent_obj = parent.get_object() if isinstance(parent, IndirectObject) else parent
                    t = parent_obj.get('/T')
            if t is None:
                continue

            terminal = terminals_by_name.get(str(t))
            if terminal is None:
                continue
            ft = terminal.get('/FT') or a.get('/FT')
            if ft != '/Btn':
                continue

            ap = a.get('/AP')
            if not ap:
                continue
            ap_n = ap.get('/N')
            if not ap_n:
                continue
            on_keys = [str(k)[1:] for k in ap_n.keys() if str(k) != '/Off']

            v_str = _name_str(terminal.get('/V'))
            target = '/Off'
            if v_str and v_str in on_keys:
                target = '/' + v_str
            a[NameObject('/AS')] = NameObject(target)


def fill(data: dict, template_path: str = TEMPLATE, output_path: str = None) -> str:
    """Fill template with data dict and write to output_path. Returns output path."""
    if output_path is None:
        week = data.get('week_ending', 'unknown').replace('/', '-')
        output_path = f'Certified_Payroll_{week}_FILLED.pdf'

    reader = PdfReader(template_path)
    writer = PdfWriter()
    writer.append(reader)

    # Shrink font + widen day-column fields BEFORE filling so pypdf uses the
    # updated /DA when it generates appearance streams.
    narrow_fields = set()
    for wi in range(7):
        for d in range(7):
            narrow_fields.add(_st_field(wi, d))
            narrow_fields.add(_ot_field(wi, d))
    for d in range(7):
        narrow_fields.add(f'Date.{d}')
    _set_font_size(writer, narrow_fields, pt=7)

    writer.update_page_form_field_values(writer.pages[0], build_field_map(data))

    # Post-process: sync widget /AS to parent /V for every checkbox/radio
    # so pdf.js + Drive preview render them correctly.
    _sync_checkbox_appearance_states(writer)

    with open(output_path, 'wb') as fh:
        writer.write(fh)
    return output_path


# ── CLI / test ────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    if len(sys.argv) > 1:
        with open(sys.argv[1]) as fh:
            data = json.load(fh)
        out = fill(data, output_path=sys.argv[2] if len(sys.argv) > 2 else None)
        print(f'✅  Filled → {out}')
    else:
        sample = {
            "payroll_number": "1",
            "week_ending": "04/11/2026",
            "employer": {
                "name":    "Metro Thermoplastic Corp",
                "address": "123 Industrial Ave, Brooklyn NY 11201",
                "email":   "info@metrothermo.com",
                "phone":   "718-555-0100",
                "tax_id":  "12-3456789"
            },
            "prime_contractor":      "NYC Department of Transportation",
            "contract_registration": "CR-20260001",
            "agency":                "NYC DOT",
            "agency_pin":            "840MTXXXXX",
            "project_address":       "Various Streets, Queens & Brooklyn NY",
            "project_name":          "FY2026 Thermoplastic Pavement Markings",
            "pla": False,
            "days": [
                {"label": "M", "date": "04/06"},
                {"label": "T", "date": "04/07"},
                {"label": "W", "date": "04/08"},
                {"label": "T", "date": "04/09"},
                {"label": "F", "date": "04/10"},
                {"label": "S", "date": "04/11"},
                {"label": "S", "date": "04/12"},
            ],
            "workers": [
                {
                    "name": "Carlos Rivera",
                    "address": "456 Oak Ave, Brooklyn NY 11203",
                    "ssn4": "1234",
                    "trade": "Pavement Marking Operator",
                    "journeyperson": True,
                    "st_hours": ["8","8","8","8","8","0","0"],
                    "ot_hours": ["0","0","0","0","2","0","0"],
                    "total_st": "40", "total_ot": "2",
                    "rate_st": "35.00", "rate_ot": "52.50",
                    "gross_pay": "1505.00", "net_pay": "1100.00",
                    "deductions": "405.00", "annualized_rate": "72800.00"
                },
                {
                    "name": "Jose Martinez",
                    "address": "789 Elm St, Queens NY 11373",
                    "ssn4": "5678",
                    "trade": "Pavement Marking Laborer",
                    "journeyperson": True,
                    "st_hours": ["8","8","8","8","8","0","0"],
                    "ot_hours": ["0","0","0","0","0","0","0"],
                    "total_st": "40", "total_ot": "0",
                    "rate_st": "28.50", "rate_ot": "42.75",
                    "gross_pay": "1140.00", "net_pay": "850.00",
                    "deductions": "290.00", "annualized_rate": "59280.00"
                },
                {
                    "name": "Maria Lopez",
                    "address": "321 Pine Rd, Bronx NY 10451",
                    "ssn4": "9012",
                    "trade": "Pavement Marking Laborer",
                    "journeyperson": True,
                    "st_hours": ["8","8","8","8","8","0","0"],
                    "ot_hours": ["0","0","0","0","0","0","0"],
                    "total_st": "40", "total_ot": "0",
                    "rate_st": "28.50", "rate_ot": "42.75",
                    "gross_pay": "1140.00", "net_pay": "855.00",
                    "deductions": "285.00", "annualized_rate": "59280.00"
                },
            ],
            "signatory": {
                "name": "Anthe Angelides",
                "title": "Principal",
                "date": "04/14/2026",
                "year": "2026"
            }
        }
        out = fill(sample, output_path='test_certified_payroll_filled.pdf')
        print(f'✅  Test fill → {out}')
