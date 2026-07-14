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


def fill_acroform(template_path: str, fields: dict, output_path: str) -> str:
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
    return output_path


def merge_filled(template_path: str, list_of_fields: list, output_path: str) -> str:
    """Fill the same template once per field-set and concatenate them into
    one PDF (all EUs → one doc, all Certs → one doc).

    Every source shares the template's field names, so before appending we
    rename each source's widget /T names with a unique per-source suffix
    (_suffix_widget_names) — otherwise pypdf dedupes by name and every page
    shows the LAST source's values. Same rename-then-merge recipe as
    fill_production_log.py; /AP is already baked, so NO flatten is needed."""
    from fill_production_log import _suffix_widget_names   # lazy: avoid import cycle

    final = PdfWriter()
    with tempfile.TemporaryDirectory() as td:
        for i, fields in enumerate(list_of_fields or []):
            tmp = str(Path(td) / f'{i}.pdf')
            fill_acroform(template_path, fields, tmp)   # bakes /AP for this source
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
