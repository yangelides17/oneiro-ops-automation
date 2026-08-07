"""
Generic AcroForm text-field filler using the project's proven process:
pypdf sets /V, then PyMuPDF (_appearances.regenerate_appearances) rebuilds
every text widget's /AP and locks /NeedAppearances=false so the values
render identically in every viewer — WITHOUT flattening, so the form stays
editable (the admin can tweak a value before printing).

Used by the on-demand month-end fill endpoint (fill_server.py) for the
Employee Utilization + Certificate docs. Same appearance pipeline the
Drive-watched fillers use (fill_certified_payroll, fill_signin, …).
"""
import os
import tempfile
from io import BytesIO
from pathlib import Path

from pypdf import PdfReader, PdfWriter

from _appearances import normalize_acroform_dr, regenerate_appearances

# doc_kind (from the webapp) → worker doc_type / TEMPLATE_FILES key.
DOC_KIND_TEMPLATE = {
    'EU':   'employee_utilization',
    'CERT': 'certificates',
}

# Identity caption drawn on continuation pages — see _stamp_identity.
STAMP_FONT      = 'hebo'   # Helvetica-Bold (PyMuPDF shorthand)
STAMP_FONTSIZE  = 10
STAMP_MARGIN_X  = 36       # 0.5in in from the right edge
STAMP_BASELINE_Y = 32      # ~0.44in down from the top


def _stamp_identity(pdf_path: str, text: str, *, skip_first_page: bool = True) -> None:
    """Draw a small right-aligned identity caption at the top of each page.

    The Employee Utilization form's second page is the back half of one
    long trade table: it carries no contract number, no borough, and no
    month. On paper that makes a separated page unmatchable to its own
    front page, and on upload it leaves the page classifier nothing to
    read — those pages were the ones coming back low-confidence.

    Drawn as page CONTENT rather than a form field on purpose: it can't be
    edited or cleared by an admin tweaking values before printing, and it
    survives any later flatten.

    `skip_first_page` because page 1 already prints the contract and month
    in its own header fields; stamping it again would just be clutter.
    """
    if not text:
        return
    import fitz   # local import: only the stamp path needs it

    doc = fitz.open(pdf_path)
    try:
        width = fitz.get_text_length(text, fontname=STAMP_FONT, fontsize=STAMP_FONTSIZE)
        for i, page in enumerate(doc):
            if skip_first_page and i == 0:
                continue
            r = page.rect
            page.insert_text(
                fitz.Point(r.x1 - STAMP_MARGIN_X - width, r.y0 + STAMP_BASELINE_Y),
                text, fontname=STAMP_FONT, fontsize=STAMP_FONTSIZE, color=(0, 0, 0),
            )
        # Same temp-then-atomic-replace recipe as regenerate_appearances:
        # PyMuPDF won't save over the path it has open except incrementally,
        # which is unreliable on files pypdf just wrote.
        fd, tmp = tempfile.mkstemp(dir=os.path.dirname(pdf_path) or '.', suffix='.pdf')
        os.close(fd)
        try:
            doc.save(tmp, garbage=0, deflate=True)
            doc.close()
            os.replace(tmp, pdf_path)
        except Exception:
            if os.path.exists(tmp):
                os.unlink(tmp)
            raise
    finally:
        if not doc.is_closed:
            doc.close()


def fill_acroform(template_path: str, fields: dict, output_path: str,
                  stamp: str = '') -> str:
    """Fill the template's AcroForm text fields from a {field_name: value}
    dict and write to output_path. Leaves fields editable; bakes correct
    /AP via PyMuPDF so values render everywhere."""
    reader = PdfReader(template_path)
    writer = PdfWriter()
    writer.append(reader)
    normalize_acroform_dr(writer)   # guard pypdf's /DR fallback crash

    clean = {k: ('' if v is None else str(v)) for k, v in (fields or {}).items()}
    # A field's widgets can span pages (e.g. a Contract # shown on every
    # page); update on each page so all widgets get set. pypdf logs a
    # harmless font-fallback warning here — regenerate_appearances fixes it.
    for page in writer.pages:
        try:
            writer.update_page_form_field_values(page, clean, auto_regenerate=False)
        except Exception:
            pass

    with open(output_path, 'wb') as fh:
        writer.write(fh)

    regenerate_appearances(output_path)   # PyMuPDF /AP bake + NeedAppearances=false
    # After the /AP bake, not before: regenerate_appearances rewrites the
    # whole file, so a caption drawn first would be at the mercy of that pass.
    _stamp_identity(output_path, stamp)
    return output_path


def merge_filled(template_path: str, items: list, output_path: str) -> str:
    """Fill the same template once per item and concatenate them into
    one PDF (all EUs → one doc, all Certs → one doc).

    `items` is a list of {'fields': {...}, 'stamp': '...'} — the stamp is
    per item, not per batch, since each source is a different contract and
    month. Stamping happens inside fill_acroform, before the merge, so the
    combined PDF's pages stay individually identifiable.

    Every source shares the template's field names, so before appending we
    rename each source's widget /T names with a unique per-source suffix
    (_suffix_widget_names) — otherwise pypdf dedupes by name and every page
    shows the LAST source's values. Same rename-then-merge recipe as
    fill_production_log.py; /AP is already baked, so NO flatten is needed."""
    from fill_production_log import _suffix_widget_names   # lazy: avoid import cycle

    final = PdfWriter()
    with tempfile.TemporaryDirectory() as td:
        for i, item in enumerate(items or []):
            item = item or {}
            tmp = str(Path(td) / f'{i}.pdf')
            # bakes /AP + stamps this source
            fill_acroform(template_path, item.get('fields') or {}, tmp,
                          stamp=str(item.get('stamp') or ''))
            w = PdfWriter()
            w.append(PdfReader(tmp))
            _suffix_widget_names(w, f'_d{i}')            # keep colliding names independent
            buf = BytesIO()
            w.write(buf)
            buf.seek(0)
            final.append(PdfReader(buf))
    with open(output_path, 'wb') as fh:
        final.write(fh)
    return output_path
