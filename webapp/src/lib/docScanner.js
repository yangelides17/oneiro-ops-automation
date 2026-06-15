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

// Approx uncompressed size of opencv.js — used to render a rough but
// honest progress bar. The streaming reader yields decompressed bytes,
// so this works whether or not the CDN gzips the response.
const EST_OPENCV_BYTES = 9 * 1024 * 1024

let _scannerPromise = null

// ── Load-progress pub/sub ─────────────────────────────────────
// So the scan UI can show a "Preparing scanner…" indicator + progress
// bar for the one-time OpenCV.js download instead of a frozen screen.
// phase: 'idle' | 'downloading' | 'initializing' | 'ready' | 'error'
let _progress = { phase: 'idle', loaded: 0, ratio: 0, indeterminate: false, error: '' }
const _listeners = new Set()

function setProgress(patch) {
  _progress = { ..._progress, ...patch }
  _listeners.forEach(fn => { try { fn(_progress) } catch { /* noop */ } })
}

export function getScannerProgress() { return _progress }

// Subscribe to load progress. Immediately invokes with the current
// value; returns an unsubscribe fn.
export function onScannerProgress(fn) {
  _listeners.add(fn)
  fn(_progress)
  return () => _listeners.delete(fn)
}

const log = (...a) => console.log('[docScanner]', ...a)
const warn = (...a) => console.warn('[docScanner]', ...a)

// Inject a <script> once; resolve when it loads. Rejects on error OR if
// neither load nor error fires within timeoutMs (a hung request).
function loadScript(src, timeoutMs = 90000) {
  return new Promise((resolve, reject) => {
    const existing = document.querySelector(`script[data-src="${src}"]`)
    if (existing && existing.dataset.loaded === '1') return resolve()
    const target = existing || document.createElement('script')
    let settled = false
    const timer = setTimeout(() => {
      if (settled) return
      settled = true
      warn('script load timed out after', timeoutMs, 'ms:', src)
      reject(new Error(`Timed out loading ${src}`))
    }, timeoutMs)
    const onLoad = () => {
      if (settled) return
      settled = true; clearTimeout(timer); target.dataset.loaded = '1'
      log('script loaded:', src)
      resolve()
    }
    const onErr = () => {
      if (settled) return
      settled = true; clearTimeout(timer)
      warn('script failed:', src)
      reject(new Error(`Failed to load ${src}`))
    }
    target.addEventListener('load', onLoad)
    target.addEventListener('error', onErr)
    if (!existing) {
      target.src = src
      target.async = true
      target.dataset.src = src
      log('injecting script:', src)
      document.head.appendChild(target)
    }
  })
}

// Best-effort streaming download so we can show a progress bar. Warms
// the HTTP cache so the subsequent <script> load is instant. May throw
// (CORS / network) — callers fall back to a plain script load. The
// reader yields decompressed bytes, so we gauge progress against an
// estimated uncompressed size rather than Content-Length.
async function prefetchWithProgress(url) {
  log('prefetch (with progress) start:', url)
  const res = await fetch(url, { mode: 'cors', cache: 'force-cache' })
  if (!res.ok || !res.body) throw new Error(`prefetch HTTP ${res.status}`)
  const reader = res.body.getReader()
  let loaded = 0
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    loaded += value.length
    setProgress({
      phase: 'downloading',
      loaded,
      ratio: Math.min(loaded / EST_OPENCV_BYTES, 0.99),
      indeterminate: false,
    })
  }
  log('prefetch complete:', (loaded / 1e6).toFixed(1), 'MB')
}

// OpenCV.js signals readiness differently across builds. Polling for a
// real class (cv.Mat) is the one universally reliable check.
function waitForCvReady(timeoutMs = 90000) {
  return new Promise((resolve, reject) => {
    const start = Date.now()
    const tick = () => {
      const cv = window.cv
      if (cv && typeof cv.Mat === 'function') { log('cv runtime ready'); return resolve(cv) }
      if (cv && cv.then) { cv.then(c => { window.cv = c; log('cv ready (promise)'); resolve(c) }).catch(reject); return }
      if (Date.now() - start > timeoutMs) {
        warn('cv init timed out; window.cv is', typeof cv)
        return reject(new Error('OpenCV.js initialized but the runtime never became ready'))
      }
      setTimeout(tick, 50)
    }
    tick()
  })
}

// Load OpenCV.js + jscanify (idempotent, memoized). Returns a jscanify
// scanner instance. Safe to call repeatedly / concurrently. Progress is
// reported via onScannerProgress(); on failure sets phase 'error'.
export function loadScanner() {
  if (_scannerPromise) return _scannerPromise
  _scannerPromise = (async () => {
    log('loadScanner: starting; OpenCV from', OPENCV_URL)
    // 1) Download OpenCV with a progress bar (best effort). If the CDN
    //    blocks CORS, fall back to an opaque script-tag download.
    try {
      setProgress({ phase: 'downloading', loaded: 0, ratio: 0, indeterminate: false, error: '' })
      await prefetchWithProgress(OPENCV_URL)
    } catch (e) {
      warn('progress prefetch unavailable, falling back to script download:', e.message)
      setProgress({ phase: 'downloading', indeterminate: true })
    }
    // 2) Run OpenCV (from cache if the prefetch succeeded).
    await loadScript(OPENCV_URL)
    setProgress({ phase: 'initializing', indeterminate: true })
    await waitForCvReady()
    // 3) Load jscanify (tiny) on top of the cv runtime.
    log('loading jscanify from', JSCANIFY_URL)
    await loadScript(JSCANIFY_URL)
    const JsScanify = window.jscanify
    if (typeof JsScanify !== 'function') throw new Error('jscanify loaded but global is missing')
    setProgress({ phase: 'ready', ratio: 1, indeterminate: false })
    log('loadScanner: ready')
    return new JsScanify()
  })().catch(err => {
    _scannerPromise = null   // allow a later retry
    warn('loadScanner FAILED:', err && err.message || err)
    setProgress({ phase: 'error', error: (err && err.message) || 'Scanner failed to load' })
    throw err
  })
  return _scannerPromise
}

// Fire-and-forget warm-up — start the big download early without
// blocking. Errors are surfaced via onScannerProgress (phase 'error').
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

export function canvasToJpeg(canvas, quality = 0.85) {
  return new Promise(resolve => canvas.toBlob(b => resolve(b), 'image/jpeg', quality))
}
