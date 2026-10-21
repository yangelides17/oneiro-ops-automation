import { normalizeHeic } from './imagesToPdf'

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const i = new Image()
    i.onload = () => resolve(i)
    i.onerror = reject
    i.src = src
  })
}

// Mirrors _crop_to_ink in workers/fill_signin.py — crops a PNG data URL
// to its non-transparent bounding box (with an 8-px margin) so the
// server-side pdf-lib overlay doesn't scale a bunch of empty canvas.
// Returns a fresh PNG data URL; falls back to the input on any error.
export async function cropPngToInk(dataUrl, marginPx = 8) {
  if (!dataUrl) return dataUrl
  try {
    const img = await loadImage(dataUrl)
    const c   = document.createElement('canvas')
    c.width   = img.width
    c.height  = img.height
    const ctx = c.getContext('2d')
    ctx.drawImage(img, 0, 0)
    const { data, width, height } = ctx.getImageData(0, 0, c.width, c.height)
    let minX = width, minY = height, maxX = -1, maxY = -1
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        if (data[(y * width + x) * 4 + 3] > 0) {
          if (x < minX) minX = x
          if (y < minY) minY = y
          if (x > maxX) maxX = x
          if (y > maxY) maxY = y
        }
      }
    }
    if (maxX < 0) return dataUrl  // fully transparent → pass through
    minX = Math.max(0, minX - marginPx)
    minY = Math.max(0, minY - marginPx)
    maxX = Math.min(width  - 1, maxX + marginPx)
    maxY = Math.min(height - 1, maxY + marginPx)
    const cw = maxX - minX + 1
    const ch = maxY - minY + 1
    const out = document.createElement('canvas')
    out.width  = cw
    out.height = ch
    out.getContext('2d').drawImage(c, minX, minY, cw, ch, 0, 0, cw, ch)
    return out.toDataURL('image/png')
  } catch {
    return dataUrl
  }
}

const MAX_SIDE = 1400   // plenty for a ~150pt-wide signature box

/**
 * Turns a photo / scan of a pen signature on paper into a transparent
 * PNG data URL cropped to the ink — the same shape SignaturePad
 * produces, so the CP overlay treats both identically.
 *
 * Paper and ink brightness are estimated per-image from luminance
 * percentiles (phone photos are rarely pure white), then each pixel's
 * alpha is how far it sits from paper toward ink. Faint pixels are
 * dropped outright so paper texture and shadows don't survive as a
 * grey haze or widen the crop box. Output ink is solid black.
 */
export async function signatureFromPhoto(file) {
  const normalized = await normalizeHeic(file)
  const url = URL.createObjectURL(normalized)
  let img
  try {
    img = await loadImage(url)
  } catch {
    throw new Error('Couldn’t read that image. Try a JPG or PNG.')
  } finally {
    URL.revokeObjectURL(url)
  }

  const scale = Math.min(1, MAX_SIDE / Math.max(img.width, img.height))
  const c = document.createElement('canvas')
  c.width  = Math.max(1, Math.round(img.width  * scale))
  c.height = Math.max(1, Math.round(img.height * scale))
  const ctx = c.getContext('2d')
  ctx.drawImage(img, 0, 0, c.width, c.height)
  const imgData = ctx.getImageData(0, 0, c.width, c.height)
  const px = imgData.data

  // Luminance histogram over opaque pixels (an already-transparent PNG
  // keeps its transparent areas transparent).
  const n = px.length / 4
  const lum = new Uint8Array(n)
  const hist = new Uint32Array(256)
  let counted = 0
  for (let i = 0; i < n; i++) {
    const o = i * 4
    const l = Math.round(0.299 * px[o] + 0.587 * px[o + 1] + 0.114 * px[o + 2])
    lum[i] = l
    if (px[o + 3] > 0) { hist[l]++; counted++ }
  }
  if (!counted) throw new Error('That image is empty.')
  const percentile = (p) => {
    const target = counted * p
    let acc = 0
    for (let v = 0; v < 256; v++) { acc += hist[v]; if (acc >= target) return v }
    return 255
  }
  const paper = percentile(0.60)   // signatures are mostly paper
  const ink   = percentile(0.005)
  if (paper - ink < 40) {
    throw new Error('Couldn’t find a signature in that image — use a dark pen on light paper, well lit.')
  }

  // Pixels within 30% of paper brightness are background; the rest ramp
  // linearly to fully opaque at the ink level.
  const cutoff = paper - 0.30 * (paper - ink)
  for (let i = 0; i < n; i++) {
    const o = i * 4
    const a = lum[i] >= cutoff ? 0 : Math.min(1, (cutoff - lum[i]) / (cutoff - ink))
    px[o] = px[o + 1] = px[o + 2] = 0
    px[o + 3] = a < 0.12 ? 0 : Math.round(a * (px[o + 3] / 255) * 255)
  }
  ctx.putImageData(imgData, 0, 0)
  return cropPngToInk(c.toDataURL('image/png'))
}
