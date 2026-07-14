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
