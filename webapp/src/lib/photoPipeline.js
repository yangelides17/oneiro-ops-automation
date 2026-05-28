// Photo capture pipeline helpers shared by FieldReport.jsx and the
// usePhotoUploadQueue hook. Three pieces:
//   • watermarkImage  — burns the DOT-required date/address/coords
//                       overlay onto a captured JPEG via canvas.
//   • formatWatermarkAddress — reverse-geocode response → 3 text lines.
//   • pendingPhotosDB — IndexedDB queue keyed by WO so blobs survive
//                       reloads / tab close / app background.

const DB_NAME    = 'oneiro-photo-queue'
const DB_VERSION = 1
const STORE      = 'pendingPhotos'

// ── IndexedDB queue ───────────────────────────────────────────
// Single object store, keyed by item id. Each record:
//   { id, wo_id, source, blob, filename, mime, captured_at, geo,
//     watermark_applied }
// `wo_id` is indexed so listForWO() can grab a slice cheaply.
function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION)
    req.onupgradeneeded = () => {
      const db = req.result
      if (!db.objectStoreNames.contains(STORE)) {
        const os = db.createObjectStore(STORE, { keyPath: 'id' })
        os.createIndex('wo_id', 'wo_id', { unique: false })
      }
    }
    req.onsuccess = () => resolve(req.result)
    req.onerror   = () => reject(req.error)
  })
}

async function tx(mode, fn) {
  const db = await openDB()
  return new Promise((resolve, reject) => {
    const t  = db.transaction(STORE, mode)
    const os = t.objectStore(STORE)
    let out
    Promise.resolve(fn(os)).then(v => { out = v }).catch(reject)
    t.oncomplete = () => resolve(out)
    t.onerror    = () => reject(t.error)
    t.onabort    = () => reject(t.error)
  })
}

export const pendingPhotosDB = {
  put(record) {
    return tx('readwrite', os => os.put(record))
  },
  delete(id) {
    return tx('readwrite', os => os.delete(id))
  },
  listForWO(wo_id) {
    return tx('readonly', os => new Promise((resolve, reject) => {
      const out = []
      const idx = os.index('wo_id')
      const req = idx.openCursor(IDBKeyRange.only(String(wo_id)))
      req.onsuccess = (e) => {
        const cur = e.target.result
        if (cur) { out.push(cur.value); cur.continue() } else resolve(out)
      }
      req.onerror = () => reject(req.error)
    }))
  },
}

// ── Address formatter ─────────────────────────────────────────
// Apps Script reverse-geocode response shape:
//   { address, city, state, zip, country }
// Returns the three text lines that go into the watermark. Empty
// pieces collapse so we never emit lines like ", NY" or " 11209".
export function formatWatermarkAddress(geo) {
  if (!geo) return []
  const cityStateZip = [geo.city, geo.state].filter(Boolean).join(', ')
    + (geo.zip ? ` ${geo.zip}` : '')
  return [
    geo.address || '',
    cityStateZip.trim(),
    geo.country || '',
  ].filter(s => s && s.trim().length)
}

// ── Watermark painter ─────────────────────────────────────────
// Draws a top-right text block on the source JPEG: timestamp, address
// lines, then `lat, lng`. Matches the NYC DOT reference style — white
// text with a dark shadow + outline so it stays legible on any backdrop.
// Font size scales with image width so the overlay reads the same on
// a portrait phone shot vs. a landscape DSLR sheet.
//
// ctx: { timestamp: Date, addressLines: string[], lat: number, lng: number }
export async function watermarkImage(file, ctx) {
  if (!file) return file
  const type = String(file?.type || '').toLowerCase()
  if (type !== 'image/jpeg' && type !== 'image/png') return file

  let bitmap
  try {
    bitmap = await createImageBitmap(file)
  } catch {
    return file
  }
  const { width: w, height: h } = bitmap

  const canvas = typeof OffscreenCanvas !== 'undefined'
    ? new OffscreenCanvas(w, h)
    : Object.assign(document.createElement('canvas'), { width: w, height: h })
  const c = canvas.getContext('2d')
  if (!c) { bitmap.close?.(); return file }
  c.drawImage(bitmap, 0, 0, w, h)
  bitmap.close?.()

  // Compose the text lines.
  const lines = []
  if (ctx?.timestamp) lines.push(formatTimestamp(ctx.timestamp))
  if (Array.isArray(ctx?.addressLines)) {
    ctx.addressLines.forEach(l => l && lines.push(l))
  }
  if (isFinite(ctx?.lat) && isFinite(ctx?.lng)) {
    lines.push(`${ctx.lat.toFixed(6)}, ${ctx.lng.toFixed(6)}`)
  }
  if (lines.length === 0) {
    return file
  }

  // Sizing — tuned against the reference screenshot. Long edge ~ 4000 px
  // → ~64 px font; short edge ~ 1500 px → ~24 px. Linear scale keeps the
  // overlay readable across phone resolutions.
  const longEdge = Math.max(w, h)
  const fontPx   = Math.max(20, Math.round(longEdge * 0.024))
  const padX     = Math.round(fontPx * 0.8)
  const padTop   = Math.round(fontPx * 0.6)
  const lineGap  = Math.round(fontPx * 1.25)

  c.font         = `600 ${fontPx}px -apple-system, "SF Pro", Helvetica, Arial, sans-serif`
  c.textAlign    = 'right'
  c.textBaseline = 'top'
  c.fillStyle    = '#ffffff'
  c.strokeStyle  = 'rgba(0,0,0,0.85)'
  c.lineWidth    = Math.max(2, Math.round(fontPx * 0.12))
  c.shadowColor  = 'rgba(0,0,0,0.7)'
  c.shadowBlur   = Math.round(fontPx * 0.4)
  c.shadowOffsetX = 0
  c.shadowOffsetY = Math.round(fontPx * 0.08)

  const x = w - padX
  lines.forEach((line, i) => {
    const y = padTop + i * lineGap
    // Stroke first for the outline halo, then fill the white glyph.
    c.strokeText(line, x, y)
    c.fillText(line, x, y)
  })

  // Re-encode. Keep JPEG quality high (q=0.92) because compressImage
  // downstream may further reduce it; we don't want compounding losses
  // softening the overlay.
  const blob = canvas.convertToBlob
    ? await canvas.convertToBlob({ type: 'image/jpeg', quality: 0.92 })
    : await new Promise(r => canvas.toBlob(r, 'image/jpeg', 0.92))
  if (!blob) return file

  const baseName = (file.name || 'photo').replace(/\.(png|jpg|jpeg|heic|heif|webp)$/i, '')
  return new File([blob], `${baseName}.jpg`, { type: 'image/jpeg' })
}

// "May 28, 2026 at 3:55:02 PM" — matches the DOT reference exactly.
function formatTimestamp(d) {
  const date = d instanceof Date ? d : new Date(d)
  const datePart = date.toLocaleDateString('en-US', {
    month: 'long', day: 'numeric', year: 'numeric',
  })
  const timePart = date.toLocaleTimeString('en-US', {
    hour: 'numeric', minute: '2-digit', second: '2-digit', hour12: true,
  })
  return `${datePart} at ${timePart}`
}
