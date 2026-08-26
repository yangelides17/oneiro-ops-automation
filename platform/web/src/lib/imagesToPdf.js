// Assemble cleaned scan images into a single multi-page PDF — one image
// per page, on a STANDARD US-Letter page (oriented to match the image),
// with the image fit-to-page and centered. This PDF flows into the
// EXISTING sign-in upload path (parse-upload → Claude Vision → submit →
// Drive storage), so the stored artifact stays a PDF and no server
// changes are needed.
//
// Why a fixed Letter size (NOT page = image pixels): pdf-lib page
// dimensions are POINTS (1/72"), so sizing a page to the image's pixel
// count produced a ~22"×28" page. Desktop apps auto-fit it, but Drive's
// web viewer and Chrome's in-browser print render at native size and
// show only a cropped corner. A normal Letter page prints correctly
// everywhere.

// US Letter in points.
const LETTER_SHORT = 612   //  8.5"
const LETTER_LONG  = 792   // 11"

// jpegBlobs: array of image/jpeg Blobs. Returns Uint8Array PDF bytes.
// pdf-lib is imported lazily so its ~400 KB only loads on the scan flow,
// never in the app-wide bundle.
export async function imagesToPdf(jpegBlobs) {
  const { PDFDocument } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  for (const blob of jpegBlobs) {
    const bytes = new Uint8Array(await blob.arrayBuffer())
    const img = await doc.embedJpg(bytes)

    // Match page orientation to the image so the sheet fills the page.
    const landscape = img.width >= img.height
    const pageW = landscape ? LETTER_LONG : LETTER_SHORT
    const pageH = landscape ? LETTER_SHORT : LETTER_LONG
    const page = doc.addPage([pageW, pageH])

    // Contain-fit: scale to fit within the page, preserving aspect,
    // centered (thin white margins if the scan's aspect differs slightly
    // from Letter — normal scanner behavior).
    const scale = Math.min(pageW / img.width, pageH / img.height)
    const w = img.width * scale
    const h = img.height * scale
    page.drawImage(img, { x: (pageW - w) / 2, y: (pageH - h) / 2, width: w, height: h })
  }
  return await doc.save()
}

// Convert a HEIC/HEIF capture (common on iPhones) to a JPEG Blob.
// Anything else passes through unchanged. heic2any is heavy, so it is
// imported lazily only when a HEIC file is actually encountered.
export async function normalizeHeic(file) {
  const isHeic = /image\/hei[cf]/i.test(file.type || '') ||
                 /\.(heic|heif)$/i.test(file.name || '')
  if (!isHeic) return file
  try {
    const { default: heic2any } = await import('heic2any')
    const out = await heic2any({ blob: file, toType: 'image/jpeg', quality: 0.85 })
    const blob = Array.isArray(out) ? out[0] : out
    const base = (file.name || 'photo').replace(/\.(heic|heif)$/i, '')
    return new File([blob], `${base}.jpg`, { type: 'image/jpeg' })
  } catch (err) {
    console.warn('[imagesToPdf] HEIC convert failed — passing original', err?.message || err)
    return file
  }
}
