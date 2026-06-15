// Assemble cleaned scan images into a single multi-page PDF — one image
// per page, each page sized to its image so nothing is cropped or
// letterboxed. This PDF is what flows into the EXISTING sign-in upload
// path (parse-upload → Claude Vision → submit → Drive storage), so the
// stored artifact stays a PDF and no server changes are needed.
//
// Page orientation is left as captured (portrait or landscape); the
// server's parse-upload rotation pre-pass normalizes sideways pages
// before Claude sees them, so we must NOT force orientation here.

// jpegBlobs: array of image/jpeg Blobs. Returns Uint8Array PDF bytes.
// pdf-lib is imported lazily so its ~400 KB only loads on the scan flow,
// never in the app-wide bundle.
export async function imagesToPdf(jpegBlobs) {
  const { PDFDocument } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  for (const blob of jpegBlobs) {
    const bytes = new Uint8Array(await blob.arrayBuffer())
    const img = await doc.embedJpg(bytes)
    const page = doc.addPage([img.width, img.height])
    page.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height })
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
