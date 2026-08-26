"""
Shared post-fill helper: regenerate every form field's /AP stream
using PyMuPDF's renderer, then lock in the result with
/NeedAppearances=false.

Why this exists
---------------
pypdf 4.3.1's automatic /AP generation has two real bugs that affect
every doc type we fill:

  1. Font fallback corruption — when a field's /DA references a font
     that isn't in /AcroForm/DR/Font (e.g. /ArialMT), pypdf falls back
     to writing /Helvetica into the /AP without registering a matching
     font dict.  Acrobat then resolves /Helvetica against its own
     system font, whose glyph widths differ from what pypdf computed
     when laying the text out — visible as letter-spacing artifacts
     and clipped overflow.

  2. No multi-line word-wrap — pypdf splits /V on `\\n` and draws each
     segment as a single line.  Long lines (e.g. a worker's full
     address) overflow the field's clip rect and get cut off by
     viewers that respect the rect.  Preview / pdf.js / Drive preview
     mask this by ignoring /AP and re-rendering from /V; Acrobat
     renders /AP faithfully and shows truncated text.

Fix
---
After pypdf has filled /V on every widget, we hand the PDF to PyMuPDF
(`fitz`) and call `widget.update()` on each widget.  PyMuPDF wraps
mupdf's text engine, which handles font fallback and multi-line wrap
correctly.  The regenerated /AP renders identically in every viewer
(Acrobat, Preview, pdf.js, Drive preview, mobile readers) and prints
the same way.

We then set /NeedAppearances=false on the AcroForm dict so viewers
trust the /AP we just wrote rather than regenerating again.  PDF 2.0
deprecates the regeneration flag and most modern viewers ignore it,
so baking correct /AP in is the durable, viewer-agnostic answer.

Usage
-----
Import and call `regenerate_appearances(path)` as the LAST step of
each filler, after `writer.write(...)` has produced the file:

    from _appearances import regenerate_appearances
    ...
    writer.write(out_fh)
    regenerate_appearances(output_path)
    return output_path
"""
import os
import tempfile

import fitz
from pypdf.generic import DictionaryObject, IndirectObject, NameObject


def normalize_acroform_dr(writer) -> None:
    """
    pypdf 4.3.1 crashes inside `_update_field_annotation` (writer.py:839)
    when a form field references a font not in the page's /DR:

        dr = acroform.get("/DR", {})
        dr = dr.get_object().get("/Font", DictionaryObject()).get_object()

    If /DR is missing the default is a plain `{}` (no .get_object()),
    and if /DR exists but its /Font is a plain dict (not a pypdf
    DictionaryObject), the second `.get_object()` call also blows up
    with `AttributeError: 'dict' object has no attribute 'get_object'`.

    Templates authored in newer Acrobat / Foxit builds sometimes ship
    without /DR, or with an inline /DR/Font dict that pypdf doesn't
    wrap correctly on append().  Normalize both entries so the
    fallback path survives.  Idempotent — safe to call before every
    fill.

    Call this on the PdfWriter BEFORE `update_page_form_field_values`.
    """
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


def regenerate_appearances(pdf_path: str) -> None:
    """
    Regenerate every form-field widget's /AP stream using PyMuPDF and
    set /NeedAppearances=false.  Mutates `pdf_path` in place.

    Idempotent — safe to call on an already-regenerated file.
    """
    doc = fitz.open(pdf_path)
    try:
        # 1. /AP regeneration — TEXT FIELDS ONLY.
        #
        # We deliberately skip button fields (/Btn = checkboxes, radios).
        # Buttons in our templates rely on either:
        #   (a) the template's pre-baked /AP/N "On"/"Off" state streams
        #       (sign-in checkboxes, CP J/A worker boxes), or
        #   (b) custom /AP streams we construct by hand (production log
        #       borough circles, which mupdf would otherwise overwrite
        #       with generic checkbox tick marks).
        # In both cases, what makes the right state visible is /AS at the
        # widget level — managed per-doc-type by _sync_checkbox_appearance_states
        # (CP) and _patch_borough_ap (production log).  Leaving /AP alone
        # for buttons keeps that work intact.
        #
        # The pypdf bugs we're fixing here only affect /Tx (text) fields:
        # font fallback corruption + multi-line wrap.  Text-only scope
        # is exactly the right scope.
        TEXT_FIELD_TYPES = (fitz.PDF_WIDGET_TYPE_TEXT,)
        for page in doc:
            widgets = page.widgets()
            if widgets is None:
                continue
            for widget in widgets:
                if widget.field_type not in TEXT_FIELD_TYPES:
                    continue
                try:
                    widget.update()
                except Exception:
                    # Don't let one bad widget abort the whole doc.
                    pass

        # 2. /NeedAppearances=false on the AcroForm dict.  PyMuPDF
        # exposes catalog manipulation via the xref API.
        catalog_xref = doc.pdf_catalog()
        acro = doc.xref_get_key(catalog_xref, 'AcroForm')
        # acro is (kind, value) where kind is 'xref', 'dict', or 'null'
        if acro and acro[0] == 'xref':
            acro_xref = int(acro[1].split()[0])
            doc.xref_set_key(acro_xref, 'NeedAppearances', 'false')
        elif acro and acro[0] == 'dict':
            inline = acro[1]
            if '/NeedAppearances true' in inline:
                inline = inline.replace('/NeedAppearances true',
                                        '/NeedAppearances false')
            elif '/NeedAppearances' not in inline:
                inline = inline.replace('<<', '<< /NeedAppearances false ', 1)
            doc.xref_set_key(catalog_xref, 'AcroForm', inline)

        # 3. Save to a temp file then atomically replace.  PyMuPDF
        # disallows saving over the same path the doc is open from
        # unless using incremental mode, which is unreliable on files
        # pypdf just produced (xref/object-stream layout differences).
        fd, tmp = tempfile.mkstemp(
            dir=os.path.dirname(pdf_path) or '.',
            suffix='.pdf',
        )
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
