// In-browser document scanner: turns a phone photo of a paper sign-in
// sheet into a clean, de-skewed, high-contrast "scanned" page — so the
// PDF we store (and send to primes with certified payroll) looks
// professional rather than like a snapshot on a desk.
//
// Built on jscanify (corner detection + perspective warp) which runs on
// OpenCV.js (wasm). Both are loaded LAZILY and only on the scan flow —
// the ~8 MB OpenCV.js never touches app startup or any other page. The
// browser caches it immutably, so it downloads once per device. Call
// prefetchScanner() when the crew opens the Sign-In page so the wasm is
// usually ready by the time they tap "Take Photos".
//
// Everything degrades gracefully: if OpenCV can't load or detection
// fails, the caller can still fall back to the raw photo so a scan is
// never blocked — it just won't be auto-cropped.

// Pinned CDN builds (jscanify's own demo pins this OpenCV release). To
// self-host for tighter cache control, drop these files under
// webapp/public/vendor/ and point these constants at /vendor/... .
const OPENCV_URL   = 'https://docs.opencv.org/4.7.0/opencv.js'
const JSCANIFY_URL = 'https://cdn.jsdelivr.net/npm/jscanify@1.2.0/src/jscanify.min.js'

let _scannerPromise = null

// Inject a <script> once; resolve when it has loaded.
function loadScript(src) {
  return new Promise((resolve, reject) => {
    const existing = document.querySelector(`script[data-src="${src}"]`)
    if (existing) {
      if (existing.dataset.loaded === '1') return resolve()
      existing.addEventListener('load', () => resolve())
      existing.addEventListener('error', () => reject(new Error(`Failed to load ${src}`)))
      return
    }
    const s = document.createElement('script')
    s.src = src
    s.async = true
    s.dataset.src = src
    s.addEventListener('load', () => { s.dataset.loaded = '1'; resolve() })
    s.addEventListener('error', () => reject(new Error(`Failed to load ${src}`)))
    document.head.appendChild(s)
  })
}

// OpenCV.js signals readiness differently across builds. Polling for a
// real class (cv.Mat) is the one universally reliable check.
function waitForCvReady(timeoutMs = 60000) {
  return new Promise((resolve, reject) => {
    const start = Date.now()
    const tick = () => {
      const cv = window.cv
      if (cv && typeof cv.Mat === 'function') return resolve(cv)
      if (cv && cv.then) { cv.then(c => { window.cv = c; resolve(c) }).catch(reject); return }
      if (Date.now() - start > timeoutMs) return reject(new Error('OpenCV.js init timed out'))
      setTimeout(tick, 40)
    }
    tick()
  })
}

// Load OpenCV.js + jscanify (idempotent, memoized). Returns a jscanify
// scanner instance. Safe to call repeatedly / concurrently.
export function loadScanner() {
  if (_scannerPromise) return _scannerPromise
  _scannerPromise = (async () => {
    await loadScript(OPENCV_URL)
    await waitForCvReady()
    await loadScript(JSCANIFY_URL)
    const JsScanify = window.jscanify
    if (typeof JsScanify !== 'function') throw new Error('jscanify failed to load')
    return new JsScanify()
  })().catch(err => {
    _scannerPromise = null   // allow a later retry
    throw err
  })
  return _scannerPromise
}

// Fire-and-forget warm-up — start the big download early without
// blocking. Swallows errors (a failed prefetch just means the first
// scan pays the cost itself).
export function prefetchScanner() {
  loadScanner().catch(() => {})
}

// Decode any image File/Blob to a canvas, downscaled so OpenCV stays
// within mobile memory limits. Mirrors photoPipeline's createImageBitmap
// + iOS <img> fallback. Returns { canvas, width, height }.
export async function imageToCanvas(file, maxEdge = 1800) {
  let drawable, w, h, cleanup
  try {
    const bitmap = await createImageBitmap(file)
    drawable = bitmap; w = bitmap.width; h = bitmap.height
    cleanup = () => bitmap.close?.()
  } catch {
    const url = URL.createObjectURL(file)
    try {
      const img = await new Promise((res, rej) => {
        const i = new Image()
        i.onload = () => res(i); i.onerror = () => rej(new Error('image decode failed'))
        i.src = url
      })
      drawable = img; w = img.naturalWidth; h = img.naturalHeight
    } finally { /* url revoked below after draw */ }
    cleanup = () => URL.revokeObjectURL(url)
  }
  const scale = Math.max(w, h) > maxEdge ? maxEdge / Math.max(w, h) : 1
  const cw = Math.round(w * scale), ch = Math.round(h * scale)
  const canvas = document.createElement('canvas')
  canvas.width = cw; canvas.height = ch
  canvas.getContext('2d').drawImage(drawable, 0, 0, cw, ch)
  cleanup?.()
  return { canvas, width: cw, height: ch }
}

// Auto-detect the sheet's four corners in a canvas. Returns
// { topLeftCorner, topRightCorner, bottomLeftCorner, bottomRightCorner }
// (jscanify's shape, each { x, y }), or a full-frame rectangle fallback
// if detection fails so the user always has draggable handles to adjust.
export async function detectCorners(canvas) {
  const fallback = {
    topLeftCorner:     { x: 0,            y: 0 },
    topRightCorner:    { x: canvas.width, y: 0 },
    bottomLeftCorner:  { x: 0,            y: canvas.height },
    bottomRightCorner: { x: canvas.width, y: canvas.height },
  }
  let src = null
  try {
    const scanner = await loadScanner()
    src = window.cv.imread(canvas)
    const contour = scanner.findPaperContour(src)
    if (!contour) return fallback
    const pts = scanner.getCornerPoints(contour)
    // Guard against degenerate detections (all-zero / collapsed quads).
    const ok = pts && ['topLeftCorner', 'topRightCorner', 'bottomLeftCorner', 'bottomRightCorner']
      .every(k => pts[k] && isFinite(pts[k].x) && isFinite(pts[k].y))
    return ok ? pts : fallback
  } catch {
    return fallback
  } finally {
    try { src?.delete?.() } catch { /* noop */ }
  }
}

// Perspective-warp the canvas to the given corners and clean it up to a
// crisp grayscale "scanned document" look. Returns a JPEG Blob. On any
// OpenCV error, falls back to re-encoding the source canvas as JPEG so a
// scan is never blocked.
export async function warpAndClean(canvas, corners) {
  try {
    const cv = window.cv || (await loadScanner(), window.cv)
    const scanner = await loadScanner()

    // Output size = average of the detected edge lengths, so the warped
    // page keeps the sheet's real aspect ratio.
    const dist = (a, b) => Math.hypot(a.x - b.x, a.y - b.y)
    const { topLeftCorner: tl, topRightCorner: tr, bottomLeftCorner: bl, bottomRightCorner: br } = corners
    const outW = Math.round((dist(tl, tr) + dist(bl, br)) / 2) || canvas.width
    const outH = Math.round((dist(tl, bl) + dist(tr, br)) / 2) || canvas.height

    const warped = scanner.extractPaper(canvas, outW, outH, corners)

    // Grayscale + adaptive threshold → the classic scanner look. Reads
    // cleanly when photographed under uneven lighting.
    const src = cv.imread(warped)
    const gray = new cv.Mat()
    cv.cvtColor(src, gray, cv.COLOR_RGBA2GRAY)
    const clean = new cv.Mat()
    cv.adaptiveThreshold(gray, clean, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C,
      cv.THRESH_BINARY, 21, 15)
    const out = document.createElement('canvas')
    out.width = clean.cols; out.height = clean.rows
    cv.imshow(out, clean)
    src.delete(); gray.delete(); clean.delete()
    return await canvasToJpeg(out)
  } catch (err) {
    console.warn('[docScanner] warpAndClean failed — using raw frame', err?.message || err)
    return await canvasToJpeg(canvas)
  }
}

function canvasToJpeg(canvas, quality = 0.85) {
  return new Promise(resolve => canvas.toBlob(b => resolve(b), 'image/jpeg', quality))
}
