/**
 * pdfMerge — combine a contract-borough's month of documents into one PDF.
 *
 * Used by the Combined-PDF download type (payroll period → month). The zip
 * gets one of these per contract-borough instead of the individual files.
 *
 * Two things about the inputs drive the design:
 *
 * 1. The archive is heterogeneous and the doc type does NOT tell you which
 *    shape you have. Measured against real archived files: today's Sign-Ins,
 *    Certified Payrolls, Employee Utilization forms and Certificates are all
 *    flat scans (0 widgets, 0 extractable text, one image per page) because
 *    they're printed, wet-signed and scanned back in. But a generated Sign-In
 *    or an untouched generated CP is a live AcroForm, and a wet-signed CP
 *    reupload is byte-different yet metadata-identical to the generated one.
 *    So we probe each source's bytes rather than trusting its type. The same
 *    path then handles electronically-signed documents when those arrive.
 *
 * 2. Field-name collision — the bug prepPdfForDelivery exists to prevent,
 *    where merging documents off a shared template collapses duplicate field
 *    names and every doc but the first renders blank — cannot occur here.
 *    PDFDocument.create() + copyPages produces a document with NO /AcroForm
 *    catalog at all (verified: the '/AcroForm' literal is absent from the
 *    output bytes), so there is no field tree for names to collapse into, and
 *    values render from the /AP streams on the widget annotations. Merging two
 *    forms with identical field names and different values keeps both intact.
 *    We still uniquify the names as cheap insurance for any future path that
 *    rebuilds the form tree.
 *
 * The rename walks page /Annots directly rather than going through
 * getForm().getFields(). That is deliberate: pypdf's writer.append() — the
 * idiom every worker filler uses — produces files whose /AcroForm/Fields
 * pdf-lib cannot walk, so getFields() returns 0 on a document that visibly
 * has 47 widgets. Renaming through the form API would silently do nothing.
 * This mirrors workers/fill_production_log.py:_suffix_widget_names.
 *
 * Never flatten, and always save with updateFieldAppearances:false — the
 * default path regenerates /AP with pdf-lib's own font logic and would
 * clobber the PyMuPDF-rendered appearances the generated docs depend on.
 */

// Cheap, conservative "could this have a form?" probe — same rule as
// server.js:canHaveAcroForm. False is provable; true is a maybe.
export function canHaveAcroForm(buf) {
  return buf.includes('/AcroForm') || buf.includes('/ObjStm')
}

const PAGE_W = 612
const PAGE_H = 792
const MARGIN = 54
const LINE_H = 15
const HEADER_H = 118          // title + prime + month + rule on the first cover page

/**
 * Suffix every widget annotation's /T on every page, following /Parent
 * chains (radio/checkbox groups keep the name on the parent). Touches only
 * /T — never /V, never /AP — so appearances are unaffected.
 * Returns the number of names changed.
 */
function suffixWidgetNames(doc, suffix, pdfLib) {
  const { PDFName, PDFString, PDFHexString, PDFDict, PDFArray } = pdfLib
  const context = doc.context
  const seen = new Set()
  let renamed = 0

  const rename = (dict, depth) => {
    if (!dict || depth > 8) return
    const key = dict.toString?.().length ? dict : null
    if (key && seen.has(dict)) return
    seen.add(dict)
    const t = dict.get(PDFName.of('T'))
    if (t instanceof PDFString || t instanceof PDFHexString) {
      dict.set(PDFName.of('T'), PDFString.of(t.decodeText() + suffix))
      renamed++
    }
    const parent = dict.get(PDFName.of('Parent'))
    if (parent) {
      try { rename(context.lookup(parent, PDFDict), depth + 1) } catch { /* orphan ref */ }
    }
  }

  for (const page of doc.getPages()) {
    let annots
    try { annots = page.node.Annots() } catch { annots = null }
    if (!annots) continue
    for (let i = 0; i < annots.size(); i++) {
      let dict
      try { dict = annots.lookup(i, PDFDict) } catch { continue }
      if (!dict) continue
      if (dict.get(PDFName.of('Subtype'))?.asString?.() !== '/Widget') continue
      rename(dict, 0)
    }
    void PDFArray
  }
  return renamed
}

const ellipsize = (font, text, size, maxW) => {
  let s = String(text == null ? '' : text)
  if (font.widthOfTextAtSize(s, size) <= maxW) return s
  while (s.length > 1 && font.widthOfTextAtSize(s + '…', size) > maxW) s = s.slice(0, -1)
  return s + '…'
}

/**
 * How many cover pages an index of `n` rows needs. The first carries the
 * header block, so it holds fewer rows than a spill page.
 */
function coverPageCount(n) {
  const first = Math.floor((PAGE_H - MARGIN - HEADER_H - MARGIN) / LINE_H)
  const spill = Math.floor((PAGE_H - MARGIN - MARGIN) / LINE_H)
  if (n <= first) return 1
  return 1 + Math.ceil((n - first) / spill)
}

/**
 * Merge `sources` into a single PDF with a cover-page index.
 *
 *   sources — [{ label, bytes }] in the order they should appear
 *   meta    — { title, contractor, month, incompleteNote }
 *
 * Returns { bytes, index, omitted, renamed, coverPages, pageCount }.
 * `omitted` lists sources that could not be parsed; they are named on the
 * cover page rather than silently dropped.
 */
export async function mergeContractPackage(sources, meta = {}) {
  const pdfLib = await import('pdf-lib')
  const { PDFDocument, StandardFonts, rgb } = pdfLib

  const out = await PDFDocument.create()
  const index = []
  const omitted = []
  let renamed = 0

  for (const [i, s] of (sources || []).entries()) {
    try {
      const src = await PDFDocument.load(s.bytes, {
        ignoreEncryption: true, updateMetadata: false,
      })
      if (canHaveAcroForm(s.bytes)) {
        try { renamed += suffixWidgetNames(src, `__d${i}`, pdfLib) } catch { /* non-fatal */ }
      }
      const pageCount = src.getPageCount()
      const copied = await out.copyPages(src, src.getPageIndices())
      copied.forEach(p => out.addPage(p))
      index.push({ label: s.label, pages: pageCount })
    } catch (e) {
      // One unreadable source must not cost the whole package — the admin
      // would have to rebuild the month to get the other 30 documents.
      omitted.push({ label: s.label, error: e.message })
    }
  }

  const bodyPages = index.reduce((a, e) => a + e.pages, 0)
  if (out.getPageCount() !== bodyPages) {
    throw new Error(`merge page-count mismatch: copied ${out.getPageCount()}, expected ${bodyPages}`)
  }

  // ── Cover page(s) ─────────────────────────────────────────────
  const rows = index.length + (omitted.length ? omitted.length + 1 : 0)
  const coverPages = coverPageCount(rows)

  // Start pages are known only once the cover length is, so resolve them here.
  let cursor = coverPages + 1
  for (const e of index) { e.startPage = cursor; cursor += e.pages }

  const font = await out.embedFont(StandardFonts.Helvetica)
  const bold = await out.embedFont(StandardFonts.HelveticaBold)
  const ink = rgb(0.08, 0.10, 0.16)
  const soft = rgb(0.45, 0.48, 0.55)

  // Insert in reverse so page 0 ends up first.
  const covers = []
  for (let i = 0; i < coverPages; i++) covers.push(out.insertPage(i, [PAGE_W, PAGE_H]))

  let page = covers[0]
  let ci = 0
  let y = PAGE_H - MARGIN

  const title = meta.title || 'Oneiro Month End Documents'
  page.drawText(ellipsize(bold, title, 15, PAGE_W - MARGIN * 2), {
    x: MARGIN, y: y - 15, size: 15, font: bold, color: ink,
  })
  y -= 38
  page.drawText(ellipsize(font, `Prime Contractor: ${meta.contractor || '—'}`, 11, PAGE_W - MARGIN * 2), {
    x: MARGIN, y, size: 11, font, color: ink,
  })
  y -= 17
  if (meta.month) {
    page.drawText(String(meta.month), { x: MARGIN, y, size: 11, font, color: soft })
    y -= 17
  }
  if (meta.incompleteNote) {
    y -= 4
    page.drawText(ellipsize(bold, meta.incompleteNote, 9.5, PAGE_W - MARGIN * 2), {
      x: MARGIN, y, size: 9.5, font: bold, color: rgb(0.65, 0.20, 0.06),
    })
    y -= 17
  }
  y -= 8
  page.drawLine({
    start: { x: MARGIN, y }, end: { x: PAGE_W - MARGIN, y },
    thickness: 0.75, color: soft,
  })
  y -= 20

  const pageNumX = PAGE_W - MARGIN
  const drawRow = (label, pageNo, { muted = false } = {}) => {
    if (y < MARGIN + LINE_H) {
      ci++
      page = covers[Math.min(ci, covers.length - 1)]
      y = PAGE_H - MARGIN
    }
    const num = pageNo == null ? '—' : String(pageNo)
    const numW = font.widthOfTextAtSize(num, 10)
    const labelMax = PAGE_W - MARGIN * 2 - numW - 18
    page.drawText(ellipsize(font, label, 10, labelMax), {
      x: MARGIN, y, size: 10, font, color: muted ? soft : ink,
    })
    page.drawText(num, { x: pageNumX - numW, y, size: 10, font, color: muted ? soft : ink })
    y -= LINE_H
  }

  for (const e of index) drawRow(e.label, e.startPage)
  if (omitted.length) {
    y -= 6
    drawRow('Could not be included:', null, { muted: true })
    for (const o of omitted) drawRow(`   ${o.label}`, null, { muted: true })
  }

  const bytes = Buffer.from(await out.save({ updateFieldAppearances: false }))

  const expected = coverPages + bodyPages
  if (out.getPageCount() !== expected) {
    throw new Error(`merge page-count mismatch after cover: ${out.getPageCount()} vs ${expected}`)
  }

  return { bytes, index, omitted, renamed, coverPages, pageCount: out.getPageCount() }
}
