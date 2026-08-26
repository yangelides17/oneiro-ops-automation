"""
fill_signin.py — Employee Daily Sign-In Log filler.

Fills the "Employee Daily Sign In Log Template_FORM.pdf" with:
  - Header fields (prime contractor, subcontractor, contract#, address, agency,
    project name, date)
  - Up to 12 employee rows (name, classification, time in, time out)
  - Crew-leader block (contractor name, title, date)
  - Signature images overlaid at the form fields' /Rect positions

Signatures are delivered inline as base64 PNG data URLs in the `data` dict.
We decode them in memory (never persisted to Drive), crop to ink bounding box,
and draw into an overlay PDF that gets merged onto the filled template.

Input schema (from Apps Script generateSignInJson_):
    {
      "_type": "signin",
      "wo_id":              "PT-11930",
      "date":               "8/16/25",
      "prime_contractor":   "Metro Express Services",
      "subcontractor":      "Oneiro Collection LLC",
      "contract_number":    "84122MBTP496 - Brooklyn",   # borough appended
      "address":            "54-35 48th St, Maspeth NY 11378",
      "agency":             "DOT",
      "project_name":       "PT-11930 | Atlantic Ave",
      "crew": [
        {
          "name":           "Carlos Solorzano",
          "classification": "SAT",
          "time_in":        "8:15 AM",
          "time_out":       "1:15 PM",
          "sig_in_b64":     "data:image/png;base64,..." | "" | null,
          "sig_out_b64":    "data:image/png;base64,..." | "" | null
        },
        ...
      ],
      "contractor_name":         "Stamati Angelides",   # crew-leader printed name
      "contractor_title":        "Crew Leader",
      "date_signed":             "8/16/25",
      "contractor_signature_b64":"data:image/png;base64,..." | "" | null
    }
"""
from __future__ import annotations

import base64
import logging
import warnings
from io import BytesIO
from pathlib import Path

logging.getLogger('pypdf').setLevel(logging.ERROR)
warnings.filterwarnings('ignore', message='.*Font dictionary.*not found.*', module='pypdf')

from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, DictionaryObject, IndirectObject

from _appearances import regenerate_appearances
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as rl_canvas
from PIL import Image


# Each employee row has 4 widgets in this order:
# (Time In, Signature In, Time Out, Signature Out)
# Field numbering matches the template's AcroForm (not sequential — PDF field
# numbers reflect creation order, e.g. row 11 is 70,72,73,71).
ROW_FIELDS = [
    ('page0_field9',  'page0_field12', 'page0_field10', 'page0_field13'),  # row 1
    ('page0_field15', 'page0_field16', 'page0_field17', 'page0_field19'),  # row 2
    ('page0_field22', 'page0_field23', 'page0_field24', 'page0_field25'),  # row 3
    ('page0_field28', 'page0_field29', 'page0_field30', 'page0_field31'),  # row 4
    ('page0_field34', 'page0_field35', 'page0_field36', 'page0_field37'),  # row 5
    ('page0_field40', 'page0_field41', 'page0_field42', 'page0_field43'),  # row 6
    ('page0_field46', 'page0_field47', 'page0_field48', 'page0_field49'),  # row 7
    ('page0_field52', 'page0_field53', 'page0_field54', 'page0_field55'),  # row 8
    ('page0_field58', 'page0_field59', 'page0_field60', 'page0_field61'),  # row 9
    ('page0_field64', 'page0_field65', 'page0_field66', 'page0_field67'),  # row 10
    ('page0_field70', 'page0_field72', 'page0_field73', 'page0_field71'),  # row 11
    ('page0_field76', 'page0_field77', 'page0_field78', 'page0_field79'),  # row 12
]
MAX_ROWS = len(ROW_FIELDS)


# Signature overlay tuning — expand target box beyond field rect so sig fills
# more space without looking cramped. Aspect ratio is preserved.
#   H_PAD = horizontal extension each side (points)
#   V_PAD = vertical   extension each side (points)
EMP_SIG_H_PAD      = 1.0
EMP_SIG_V_PAD      = 4.5   # employee-row sig rects are ~19pt tall → target ~28pt
LEADER_SIG_H_PAD   = 1.0
LEADER_SIG_V_PAD   = 3.0   # crew-leader sig rect is ~24pt tall


def _decode_data_url(b64_or_data_url: str | None) -> bytes | None:
    """Accept raw base64 or a data URL (data:image/png;base64,...). Return
    PNG bytes, or None if the input is empty/falsy."""
    if not b64_or_data_url:
        return None
    s = b64_or_data_url
    if s.startswith('data:'):
        _, _, s = s.partition(',')
    try:
        return base64.b64decode(s, validate=False)
    except Exception:
        return None


def _crop_to_ink(png_bytes: bytes, margin_px: int = 8) -> ImageReader | None:
    """Crop a PNG to its non-transparent bounding box (with a little margin)
    so the sig fills its target box instead of sitting inside empty canvas."""
    try:
        im = Image.open(BytesIO(png_bytes)).convert('RGBA')
    except Exception:
        return None
    bbox = im.getbbox()
    if bbox is None:  # fully transparent
        return None
    x0, y0, x1, y1 = bbox
    x0 = max(0, x0 - margin_px)
    y0 = max(0, y0 - margin_px)
    x1 = min(im.width,  x1 + margin_px)
    y1 = min(im.height, y1 + margin_px)
    cropped = im.crop((x0, y0, x1, y1))
    buf = BytesIO()
    cropped.save(buf, format='PNG')
    buf.seek(0)
    return ImageReader(buf)


def _field_rects(template_path: str) -> dict:
    """Return {field_name: (page_idx, [x1, y1, x2, y2])} for every widget
    annotation on every page. Used to get the exact rect for each signature
    field so the overlay lines up with the filled form field underneath."""
    reader = PdfReader(template_path)
    rects = {}
    for pg_idx, page in enumerate(reader.pages):
        annots = page.get('/Annots') or []
        for a in annots:
            obj = a.get_object() if hasattr(a, 'get_object') else a
            if obj.get('/Subtype') != '/Widget':
                continue
            name = obj.get('/T')
            parent = obj.get('/Parent')
            if not name and parent:
                name = parent.get_object().get('/T')
            if not name:
                continue
            rect = obj.get('/Rect')
            if rect is None:
                continue
            rects[str(name)] = (pg_idx, [float(v) for v in rect])
    return rects


# Project Name / Location overlay tuning. The AcroForm field on the
# template is a single-line text box that silently clips when the
# string is too long for its visible rect — fine for one WO, useless
# for multi-WO sign-ins where the field carries every "<WO> | <Loc>"
# pair. Instead of relying on the field's appearance stream, we draw
# the text ourselves on the overlay PDF: starting at the field's
# bottom-left corner and wrapping into the whitespace beneath until
# we hit the table headers. Width spans nearly the full page width
# (we have nothing to the field's right except the Date label, which
# sits one row UP).
PROJECT_NAME_FONT_NAME    = 'Helvetica'
PROJECT_NAME_FONT_MAX     = 9      # try 9pt first; shrink as needed
PROJECT_NAME_FONT_MIN     = 6      # below 6pt readability falls off a cliff —
                                   # if the text still doesn't fit at 6pt
                                   # we ellipsize and signal a follow-up
PROJECT_NAME_LINE_RATIO   = 1.5    # line height = font_size × ratio. 1.5×
                                   # keeps line 2 clear of the printed
                                   # underline that runs across the field
                                   # row, instead of clipping into it.
PROJECT_NAME_MAX_LINES    = 2      # both lines stay within the Project
                                   # Name/Location row's left half (i.e.,
                                   # never extend under the Date field area)
                                   # so the layout stays unambiguous
PROJECT_NAME_DATE_GAP     = 28     # pt gap between text-wrap edge and the
                                   # Date_Header field's left edge — wide
                                   # enough that even Helvetica's bolder
                                   # glyphs don't crowd the "Date" label
                                   # next to the field


def _wrap_project_name(canvas, text: str, line_widths: list[float],
                       font_name: str, font_size: float) -> list[str]:
    """Greedy wrap with VARIABLE per-line widths. line_widths is a list
    of pt widths — line N uses line_widths[min(N, len-1)]. This lets us
    keep line 1 narrow (must stop short of the Date field on the same
    row) while lines 2+ can extend to the full page width since they
    flow into empty space below.

    Tries to keep '<WO> | <Loc>; ' tokens atomic so a pair never gets
    split across lines, and falls back to word-level wrap only if a
    single token is wider than the available width."""
    if not text:
        return []
    parts = text.split('; ')
    tokens = [p + ('; ' if i < len(parts) - 1 else '') for i, p in enumerate(parts)]

    def width_for(idx: int) -> float:
        return line_widths[min(idx, len(line_widths) - 1)] if line_widths else 0.0

    lines: list[str] = []
    current = ''
    for tok in tokens:
        idx = len(lines)
        max_w = width_for(idx)
        cand = current + tok
        if canvas.stringWidth(cand, font_name, font_size) <= max_w:
            current = cand
            continue
        if current:
            lines.append(current.rstrip())
            current = ''
            idx = len(lines)
            max_w = width_for(idx)
        # Token alone — try by spaces if even it's too wide.
        if canvas.stringWidth(tok, font_name, font_size) <= max_w:
            current = tok
        else:
            words = tok.split(' ')
            buf = ''
            for w in words:
                idx2 = len(lines)
                max_w2 = width_for(idx2)
                trial = (buf + (' ' if buf else '') + w)
                if canvas.stringWidth(trial, font_name, font_size) <= max_w2:
                    buf = trial
                else:
                    if buf:
                        lines.append(buf)
                    buf = w
            current = buf
    if current:
        lines.append(current.rstrip())
    return lines


def _build_overlay(crew: list, crew_leader_sig_b64: str | None,
                   rects: dict, page_size: tuple,
                   project_name: str = '') -> BytesIO | None:
    """Build a one-page PDF with signature PNGs drawn at each field rect
    (expanded for breathing room) plus the Project Name text wrapped
    into the whitespace below the field. Returns None if nothing to draw."""
    page_w, page_h = page_size
    buf = BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(page_w, page_h))
    drew_anything = False

    # ── Project Name / Location wrapped overlay ──────────────────
    if project_name and 'Project_Name' in rects:
        _pg, (px1, py1, px2, _py2) = rects['Project_Name']

        # Both lines stop short of the Date field's left edge — even
        # the second line (which sits in the empty row below) is kept
        # narrow on purpose so the rendered text feels confined to the
        # Project Name / Location half of the form rather than
        # ambiguously spilling under the Date area.
        date_rect = rects.get('Date_Header')
        if date_rect:
            _date_pg, (date_x1, _, _, _) = date_rect
            right_x = date_x1 - PROJECT_NAME_DATE_GAP
        else:
            right_x = page_w * 0.6
        narrow_w = max(60.0, right_x - px1)
        line_widths = [narrow_w] * PROJECT_NAME_MAX_LINES

        # Adaptive font sizing. Walk PROJECT_NAME_FONT_MAX → MIN, picking
        # the largest size that wraps the text into ≤ MAX_LINES. Falls
        # through to MIN with ellipsis if even the smallest still
        # overflows (signals to the caller / form-reader that text was
        # truncated rather than silently dropping it).
        chosen_size = PROJECT_NAME_FONT_MIN
        chosen_lines: list[str] = []
        for size in range(PROJECT_NAME_FONT_MAX, PROJECT_NAME_FONT_MIN - 1, -1):
            c.setFont(PROJECT_NAME_FONT_NAME, size)
            wrapped = _wrap_project_name(
                c, project_name, line_widths, PROJECT_NAME_FONT_NAME, size,
            )
            if len(wrapped) <= PROJECT_NAME_MAX_LINES:
                chosen_size, chosen_lines = size, wrapped
                break
        else:
            # Fell through — overflow at MIN size; ellipsize.
            c.setFont(PROJECT_NAME_FONT_NAME, PROJECT_NAME_FONT_MIN)
            wrapped = _wrap_project_name(
                c, project_name, line_widths,
                PROJECT_NAME_FONT_NAME, PROJECT_NAME_FONT_MIN,
            )
            wrapped = wrapped[:PROJECT_NAME_MAX_LINES]
            last = wrapped[-1].rstrip(',; ')
            while c.stringWidth(last + '…', PROJECT_NAME_FONT_NAME,
                                PROJECT_NAME_FONT_MIN) > narrow_w and last:
                last = last[:-1]
            wrapped[-1] = last + '…'
            chosen_size, chosen_lines = PROJECT_NAME_FONT_MIN, wrapped

        c.setFont(PROJECT_NAME_FONT_NAME, chosen_size)
        line_height = chosen_size * PROJECT_NAME_LINE_RATIO
        # Nudge the first line a couple points above the field bottom so
        # its descenders stay clear of the printed underline beneath.
        baseline_y = py1 + 2
        for i, line in enumerate(chosen_lines):
            c.drawString(px1, baseline_y - i * line_height, line)
        drew_anything = True

    def draw_at(field_name: str, b64: str | None, h_pad: float, v_pad: float):
        nonlocal drew_anything
        if field_name not in rects:
            return
        png = _decode_data_url(b64)
        if png is None:
            return
        img = _crop_to_ink(png)
        if img is None:
            return
        _pg, (x1, y1, x2, y2) = rects[field_name]
        rx = x1 - h_pad
        ry = y1 - v_pad
        rw = (x2 - x1) + 2 * h_pad
        rh = (y2 - y1) + 2 * v_pad
        c.drawImage(img, rx, ry, width=rw, height=rh,
                    mask='auto', preserveAspectRatio=True, anchor='c')
        drew_anything = True

    for i, emp in enumerate(crew[:MAX_ROWS]):
        _ti, sig_in, _to, sig_out = ROW_FIELDS[i]
        draw_at(sig_in,  emp.get('sig_in_b64'),  EMP_SIG_H_PAD, EMP_SIG_V_PAD)
        draw_at(sig_out, emp.get('sig_out_b64'), EMP_SIG_H_PAD, EMP_SIG_V_PAD)

    draw_at('Contractor_Signature', crew_leader_sig_b64,
            LEADER_SIG_H_PAD, LEADER_SIG_V_PAD)

    if not drew_anything:
        return None
    c.showPage()
    c.save()
    buf.seek(0)
    return buf


def _ensure_acroform_dr(writer: PdfWriter):
    """pypdf 4.3.1 crashes inside _update_field_annotation when a form
    field references a font not in the page's /DR — the fallback path
    expects /DR and /DR/Font to be DictionaryObject, but newer-Acrobat
    templates sometimes ship them as plain dicts (or omit /DR entirely).
    Normalize both. Same shape as the fix applied to fill_certified_payroll.
    Idempotent — safe to call before every fill."""
    acro = writer._root_object.get('/AcroForm')
    if acro is None:
        return
    acro_obj = acro.get_object() if isinstance(acro, IndirectObject) else acro
    dr = acro_obj.get('/DR')
    dr_obj = dr.get_object() if isinstance(dr, IndirectObject) else dr
    if not isinstance(dr_obj, DictionaryObject):
        dr_obj = DictionaryObject()
        acro_obj[NameObject('/DR')] = dr_obj
    font = dr_obj.get('/Font')
    font_obj = font.get_object() if isinstance(font, IndirectObject) else font
    if not isinstance(font_obj, DictionaryObject):
        dr_obj[NameObject('/Font')] = DictionaryObject()


def fill(data: dict, template_path: str, output_path: str) -> str:
    """Fill the Sign-In Log template and write to output_path. Returns the
    output path on success."""
    # ── 1. Fill text fields via pypdf form-fill ─────────────────────────
    reader = PdfReader(template_path)
    writer = PdfWriter(clone_from=reader)
    _ensure_acroform_dr(writer)

    f: dict[str, str] = {
        'Prime_Contractor':     data.get('prime_contractor', ''),
        'Subcontractor':        data.get('subcontractor', ''),
        'Contract_Number':      data.get('contract_number', ''),
        'Contractor_Address':   data.get('address', ''),
        'Agency':               data.get('agency', ''),
        # Project_Name is intentionally left empty in the AcroForm fill
        # — the actual text is rendered as a wrapped overlay (see
        # _build_overlay) so multi-WO sign-ins don't get clipped. If we
        # ALSO filled the AcroForm here, its truncated value would draw
        # underneath the overlay and show through.
        'Project_Name':         '',
        'Date_Header':          data.get('date', ''),
        'Contractor_Name':      data.get('contractor_name', ''),
        'Contractor_Title':     data.get('contractor_title', ''),
        'Date_Signature_Block': data.get('date_signed', ''),
    }

    crew = data.get('crew', []) or []
    for i, emp in enumerate(crew[:MAX_ROWS]):
        n = i + 1
        f[f'Employee_Name_{n}']  = emp.get('name', '')
        f[f'Employee_Class_{n}'] = emp.get('classification', '')
        ti, _sig_in, to, _sig_out = ROW_FIELDS[i]
        f[ti] = emp.get('time_in', '') or ''
        f[to] = emp.get('time_out', '') or ''

    writer.update_page_form_field_values(writer.pages[0], f)

    filled_buf = BytesIO()
    writer.write(filled_buf)
    filled_buf.seek(0)

    # ── 2. Build signature overlay ─────────────────────────────────────
    rects = _field_rects(template_path)
    mb = reader.pages[0].mediabox
    page_size = (float(mb.width), float(mb.height))
    overlay_buf = _build_overlay(
        crew,
        data.get('contractor_signature_b64'),
        rects,
        page_size,
        project_name=data.get('project_name', ''),
    )

    # ── 3. Merge overlay (if any) onto filled PDF ──────────────────────
    if overlay_buf is not None:
        filled_reader  = PdfReader(filled_buf)
        overlay_reader = PdfReader(overlay_buf)
        out = PdfWriter(clone_from=filled_reader)
        out.pages[0].merge_page(overlay_reader.pages[0])
    else:
        filled_buf.seek(0)
        out = PdfWriter(clone_from=PdfReader(filled_buf))

    with open(output_path, 'wb') as fh:
        out.write(fh)

    # Regenerate /AP via PyMuPDF + lock /NeedAppearances=false. See
    # workers/_appearances.py — applies to every doc type now.
    regenerate_appearances(output_path)
    print(f'✅  Filled Sign-In Log → {output_path}')
    return output_path


if __name__ == '__main__':
    # Smoke test with demo data (expects signatures already generated in
    # /tmp/oneiro_sig_demo/ from the earlier capture step).
    import json
    demo_sig_dir = Path('/tmp/oneiro_sig_demo')

    def _to_b64(path: Path) -> str:
        return 'data:image/png;base64,' + base64.b64encode(path.read_bytes()).decode()

    demo = {
        '_type':             'signin',
        'wo_id':             'PT-11930',
        'date':              '8/16/25',
        'prime_contractor':  'Metro Express Services',
        'subcontractor':     'Oneiro Collection LLC',
        'contract_number':   '84122MBTP496 - Brooklyn',
        'address':           '54-35 48th St, Maspeth NY 11378',
        'agency':            'DOT',
        'project_name':      'PT-11930 | Atlantic Ave',
        'contractor_name':   'Stamati Angelides',
        'contractor_title':  'Crew Leader',
        'date_signed':       '8/16/25',
        'contractor_signature_b64': _to_b64(demo_sig_dir / 'sig_stamati_crew_leader.png'),
        'crew': [
            {
                'name': 'Carlos Solorzano', 'classification': 'SAT',
                'time_in': '8:15 AM', 'time_out': '1:15 PM',
                'sig_in_b64':  _to_b64(demo_sig_dir / 'sig_carlos_in.png'),
                'sig_out_b64': _to_b64(demo_sig_dir / 'sig_carlos_out.png'),
            },
            {
                'name': 'Stamati Angelides', 'classification': 'LP',
                'time_in': '8:15 AM', 'time_out': '1:15 PM',
                'sig_in_b64':  _to_b64(demo_sig_dir / 'sig_stamati_in.png'),
                'sig_out_b64': _to_b64(demo_sig_dir / 'sig_stamati_out.png'),
            },
        ],
    }
    root = Path(__file__).resolve().parent.parent
    tpl  = root / '_template_cache' / 'Employee Daily Sign In Log Template_FORM.pdf'
    out  = Path('/tmp/oneiro_sig_demo/SignIn_PROD_FILL.pdf')
    fill(demo, str(tpl), str(out))
