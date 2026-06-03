// usePhotoUploadQueue — the brains of the Field Report photo flow.
//
// Owns the photoItems[] state for the currently-selected WO. Bridges
// three sources:
//   1. Historic photos already on Drive (GET /api/wo-photos/:woId).
//   2. Pending IndexedDB entries from prior sessions (replayed on mount).
//   3. New captures + library imports added during this session.
//
// Pipeline per item:
//   capture:  pending → geocoding → watermarking → uploading → uploaded
//   library:  pending → uploading → uploaded
//   any step can transition to `error` (with a message) — submit blocks
//   on those. The blob stays in IndexedDB until uploaded() returns 200,
//   so a mid-pipeline reload picks back up automatically.

import { useEffect, useRef, useState, useCallback } from 'react'
import {
  pendingPhotosDB,
  watermarkImage,
  formatWatermarkAddress,
} from './photoPipeline'

const newId = () => `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`

// Compress before upload — same parameters as the old submit-time
// compressor. Keeps the Drive folder lean and the upload fast.
const COMPRESS_MAX_EDGE = 2048
const COMPRESS_QUALITY  = 0.85
const COMPRESS_SKIP_BELOW_BYTES = 500 * 1024
async function compressForUpload(file) {
  const type = String(file?.type || '').toLowerCase()
  if (type !== 'image/jpeg' && type !== 'image/png') return file
  if (file.size <= COMPRESS_SKIP_BELOW_BYTES) return file
  let bitmap
  try { bitmap = await createImageBitmap(file) } catch { return file }
  const { width, height } = bitmap
  const longEdge = Math.max(width, height)
  const scale = longEdge > COMPRESS_MAX_EDGE ? COMPRESS_MAX_EDGE / longEdge : 1
  const w = Math.round(width * scale)
  const h = Math.round(height * scale)
  const canvas = typeof OffscreenCanvas !== 'undefined'
    ? new OffscreenCanvas(w, h)
    : Object.assign(document.createElement('canvas'), { width: w, height: h })
  const ctx = canvas.getContext('2d')
  if (!ctx) return file
  ctx.drawImage(bitmap, 0, 0, w, h)
  bitmap.close?.()
  const blob = canvas.convertToBlob
    ? await canvas.convertToBlob({ type: 'image/jpeg', quality: COMPRESS_QUALITY })
    : await new Promise(r => canvas.toBlob(r, 'image/jpeg', COMPRESS_QUALITY))
  if (!blob || blob.size >= file.size) return file
  const baseName = (file.name || 'photo').replace(/\.(png|jpg|jpeg|heic|heif|webp)$/i, '')
  return new File([blob], `${baseName}.jpg`, { type: 'image/jpeg' })
}

// Bounded fetch — every external call gets a hard timeout. Apps Script's
// proxied responses can stall indefinitely on a cold spreadsheet read or
// a flaky Maps service call, and the user's first symptom is "the photo
// just hangs there". Better to fail fast and let the retry path / kick
// loop handle it.
async function fetchWithTimeout(input, init, timeoutMs) {
  const ctrl = new AbortController()
  const timer = setTimeout(() => ctrl.abort(), timeoutMs)
  try {
    return await fetch(input, { ...(init || {}), signal: ctrl.signal })
  } finally {
    clearTimeout(timer)
  }
}

async function reverseGeocode(lat, lng) {
  const t0 = performance.now()
  try {
    // 6s: a healthy Apps Script Maps.newGeocoder round-trip is 1-2s.
    // Anything past 6s is a cold-start or a stuck service — we'd
    // rather watermark with just coords than hold up the upload.
    const res = await fetchWithTimeout('/api/reverse-geocode', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({ lat, lng }),
    }, 6000)
    const data = await res.json().catch(() => null)
    console.log(`[photo] reverse-geocode ${(performance.now() - t0).toFixed(0)}ms`,
      data?.error ? `error=${data.error}` : 'ok')
    if (!data || data.error) return null
    return data
  } catch (err) {
    console.warn(`[photo] reverse-geocode FAILED ${(performance.now() - t0).toFixed(0)}ms`,
      err?.message || err)
    return null
  }
}

async function uploadToDrive(file, wo_id) {
  const t0 = performance.now()
  const form = new FormData()
  form.append('photo', file)
  form.append('wo_id', wo_id)
  // 90s: covers a slow 1-2MB JPEG over a marginal LTE connection plus
  // the Apps Script side (folder resolve + base64 decode + Drive
  // create). Past 90s, it's almost always stuck — abort and let the
  // user retry. The blob stays in IndexedDB so retry is cheap.
  let res
  try {
    res = await fetchWithTimeout('/api/upload-photo', { method: 'POST', body: form }, 90000)
  } catch (err) {
    const ms = (performance.now() - t0).toFixed(0)
    console.warn(`[photo] upload FAILED ${ms}ms size=${file.size}`, err?.message || err)
    if (err?.name === 'AbortError') {
      throw new Error('Upload stalled after 90s — tap Retry')
    }
    throw err
  }
  const data = await res.json().catch(() => ({}))
  const ms = (performance.now() - t0).toFixed(0)
  if (!res.ok || data.error) {
    console.warn(`[photo] upload HTTP ${res.status} ${ms}ms size=${file.size}`, data?.error)
    throw new Error(data.error || `HTTP ${res.status}`)
  }
  console.log(`[photo] upload ok ${ms}ms size=${file.size}`)
  return data   // { success, file_id, file_url }
}

export function usePhotoUploadQueue(woId) {
  // items shape:
  //   { id, source: 'capture' | 'library' | 'historic',
  //     status: 'pending' | 'geocoding' | 'watermarking' | 'uploading' | 'uploaded' | 'error',
  //     blob?, previewUrl?, captured_at, geo?,
  //     drive_file_id?, drive_file_url?, drive_url?, thumbnail_b64?,
  //     filename, mime, error? }
  const [items, setItems]   = useState([])
  const [historicLoading, setHistoricLoading] = useState(false)
  const [historicError,   setHistoricError]   = useState(null)
  // Lock so concurrent driveItems for the same id don't double-fire.
  const inFlightRef = useRef(new Set())
  // Live ref so callbacks reading "current items" inside async work see
  // the latest snapshot instead of a stale closure.
  const itemsRef = useRef(items)
  useEffect(() => { itemsRef.current = items }, [items])

  // Patch one item by id without losing concurrent updates.
  const patch = useCallback((id, fields) => {
    setItems(prev => prev.map(it => it.id === id ? { ...it, ...fields } : it))
  }, [])
  const removeLocal = useCallback((id) => {
    setItems(prev => prev.filter(it => it.id !== id))
  }, [])

  // ── 1. Historic load on woId change ─────────────────────────
  useEffect(() => {
    setItems([])
    setHistoricError(null)
    if (!woId) return
    let cancelled = false
    setHistoricLoading(true)
    Promise.all([
      fetch(`/api/wo-photos/${encodeURIComponent(woId)}`)
        .then(r => r.json())
        .then(d => {
          if (d.error) throw new Error(d.error)
          return d.photos || []
        }),
      pendingPhotosDB.listForWO(woId).catch(() => []),
    ]).then(([historic, pending]) => {
      if (cancelled) return
      const historicItems = historic.map(p => ({
        id:             `hist-${p.file_id}`,
        source:         'historic',
        status:         'uploaded',
        drive_file_id:  p.file_id,
        drive_file_url: p.url,
        thumbnail_b64:  p.thumbnail_b64,
        mime:           p.mime,
        filename:       p.name,
        captured_at:    p.created_at,
      }))
      const pendingItems = pending.map(rec => ({
        id:                 rec.id,
        source:             rec.source || 'capture',
        status:             'pending',
        blob:               rec.blob,
        previewUrl:         rec.blob ? URL.createObjectURL(rec.blob) : null,
        captured_at:        rec.captured_at,
        geo:                rec.geo || null,
        filename:           rec.filename,
        mime:               rec.mime,
        watermark_applied:  !!rec.watermark_applied,
      }))
      setItems([...pendingItems, ...historicItems])
      // Drain anything that was sitting in IndexedDB on next tick.
      pendingItems.forEach(it => kick(it))
    }).catch(err => {
      if (!cancelled) {
        setHistoricError(err.message || 'Failed to load photos')
        setItems([])
      }
    }).finally(() => {
      if (!cancelled) setHistoricLoading(false)
    })
    return () => {
      cancelled = true
      // Revoke any object URLs we created on mount cleanup.
      itemsRef.current.forEach(it => {
        if (it.previewUrl) URL.revokeObjectURL(it.previewUrl)
      })
    }
  }, [woId])   // eslint-disable-line react-hooks/exhaustive-deps

  // ── 2. The state-machine driver ─────────────────────────────
  // `kick(item)` looks at the item's current status and advances it one
  // step. Re-entered on resume (visibilitychange, online event, manual
  // retry). Idempotent via inFlightRef.
  const kick = useCallback(async (initial) => {
    if (!initial || !initial.id) return
    if (inFlightRef.current.has(initial.id)) return
    inFlightRef.current.add(initial.id)
    try {
      // Always read the freshest copy of the item — earlier state may
      // have been patched while we were waiting in line.
      const fresh = () => itemsRef.current.find(i => i.id === initial.id)

      // Phase A: capture-source needs reverse-geocode + watermark first.
      let item = fresh()
      if (!item) return
      const tStart = performance.now()
      const timing = {}
      // Track the blob locally — Phase B must NOT re-read from itemsRef
      // because React updates the ref through a useEffect that hasn't
      // fired yet (`patch` schedules a render; the ref-sync runs only
      // after commit). Reading from itemsRef here returned the original
      // pre-watermark blob, which is why Drive received un-watermarked
      // photos while the in-memory preview correctly showed the overlay.
      let liveBlob = item.blob
      if (item.source === 'capture' && !item.watermark_applied) {
        patch(item.id, { status: 'geocoding' })
        let geoData = null
        const tGeo = performance.now()
        if (item.geo && isFinite(item.geo.lat) && isFinite(item.geo.lng)) {
          geoData = await reverseGeocode(item.geo.lat, item.geo.lng)
        }
        timing.geocodeMs = Math.round(performance.now() - tGeo)
        if (!fresh()) return
        patch(item.id, { status: 'watermarking' })
        const tWm = performance.now()
        // watermarkImage now does scale + watermark + encode at q=0.85
        // in ONE canvas pass, producing an upload-ready blob. There is
        // no second compressForUpload pass anymore — that was the old
        // double-decode/double-encode penalty.
        const watermarked = await watermarkImage(liveBlob, {
          timestamp: item.captured_at ? new Date(item.captured_at) : new Date(),
          addressLines: formatWatermarkAddress(geoData),
          lat: item.geo?.lat,
          lng: item.geo?.lng,
        })
        timing.watermarkMs = Math.round(performance.now() - tWm)
        console.log(`[photo] watermark ${timing.watermarkMs}ms in=${liveBlob?.size} out=${watermarked?.size}`)
        liveBlob = watermarked
        // ONE IndexedDB write per photo, post-watermark. The old code
        // also wrote the 3-5 MB original at addCapture time, which on
        // iOS Safari serialises Blobs through SQLite and can stall
        // 5-20 sec. We accept that a page-reload during the ~500 ms
        // watermark window loses the photo — much better than a 5-30
        // sec per-photo penalty on every capture.
        const tDb = performance.now()
        await pendingPhotosDB.put({
          id:                item.id,
          wo_id:             woId,
          source:            item.source,
          blob:              watermarked,
          filename:          item.filename,
          mime:              watermarked.type || item.mime,
          captured_at:       item.captured_at,
          geo:               item.geo,
          watermark_applied: true,
        })
        timing.dbWriteMs = Math.round(performance.now() - tDb)
        if (item.previewUrl) URL.revokeObjectURL(item.previewUrl)
        const previewUrl = URL.createObjectURL(watermarked)
        patch(item.id, { blob: watermarked, previewUrl, watermark_applied: true })
      } else if (item.source === 'library') {
        // Library imports — no watermark, still need to compress before
        // upload (camera-roll JPEGs are 3-5 MB raw) and land in
        // IndexedDB so a mid-upload reload recovers them.
        const tCompress = performance.now()
        liveBlob = await compressForUpload(liveBlob)
        timing.compressMs = Math.round(performance.now() - tCompress)
        const tDb = performance.now()
        await pendingPhotosDB.put({
          id:                item.id,
          wo_id:             woId,
          source:            item.source,
          blob:              liveBlob,
          filename:          item.filename,
          mime:              item.mime,
          captured_at:       item.captured_at,
          geo:               null,
          watermark_applied: true,
        })
        timing.dbWriteMs = Math.round(performance.now() - tDb)
      }

      // Phase B: upload to Drive. Uses liveBlob (the local var) so we
      // never race with React's state flush — even if itemsRef still
      // points at the pre-watermark blob, liveBlob is definitively the
      // bytes we want on Drive.
      if (!fresh() || !liveBlob) return
      patch(item.id, { status: 'uploading' })
      const tUpload = performance.now()
      const res = await uploadToDrive(liveBlob, woId)
      timing.uploadMs = Math.round(performance.now() - tUpload)
      const tDelete = performance.now()
      // Clear the queue row first — if we crash before the state patch
      // we'd rather lose the thumbnail than ship the photo twice.
      await pendingPhotosDB.delete(item.id).catch(() => {})
      timing.dbDeleteMs = Math.round(performance.now() - tDelete)
      timing.totalMs = Math.round(performance.now() - tStart)
      timing.shipSizeKB = Math.round((liveBlob?.size || 0) / 1024)
      timing.serverMs = res._server_ms
      timing.appsScriptMs = res._appsScript_ms
      console.log('[photo] pipeline complete', timing)
      patch(item.id, {
        status:         'uploaded',
        drive_file_id:  res.file_id,
        drive_file_url: res.file_url,
        timing,
      })
    } catch (err) {
      patch(initial.id, { status: 'error', error: err?.message || 'Upload failed' })
    } finally {
      inFlightRef.current.delete(initial.id)
    }
  }, [woId, patch])

  // ── 3. Public adders ────────────────────────────────────────
  const addCapture = useCallback((file, geo) => {
    if (!file || !woId) return
    // NOTE: deliberately NOT awaiting an IndexedDB write here. The old
    // code did `await pendingPhotosDB.put({blob: file})` with the raw
    // 3-5 MB camera JPEG, which on iOS Safari can stall 5-20 sec for a
    // single large-blob write. Now the only persistence write happens
    // post-watermark with the much smaller (~1 MB) finished blob. The
    // narrow window (~capture → watermark, typically <2 sec) where a
    // page reload could lose the photo is acceptable.
    const id = newId()
    const captured_at = new Date().toISOString()
    const item = {
      id, source: 'capture', status: 'pending', blob: file,
      previewUrl: URL.createObjectURL(file),
      captured_at, geo: geo || null,
      filename: file.name || `photo-${id}.jpg`,
      mime: file.type || 'image/jpeg',
      watermark_applied: false,
    }
    setItems(prev => [item, ...prev])
    queueMicrotask(() => kick(item))
  }, [woId, kick])

  const addLibrary = useCallback((files) => {
    if (!woId || !files || !files.length) return
    // Same rationale as addCapture: defer the IndexedDB write to kick()
    // so we don't block the UI on a slow iOS Safari blob write per file.
    const records = Array.from(files).map(f => {
      const id = newId()
      return {
        id, wo_id: woId, source: 'library', blob: f,
        filename: f.name || `library-${id}.jpg`, mime: f.type || 'image/jpeg',
        captured_at: new Date().toISOString(),
        geo: null, watermark_applied: true,
      }
    })
    const newItems = records.map(rec => ({
      id: rec.id, source: 'library', status: 'pending', blob: rec.blob,
      previewUrl: URL.createObjectURL(rec.blob),
      captured_at: rec.captured_at, geo: null,
      filename: rec.filename, mime: rec.mime, watermark_applied: true,
    }))
    setItems(prev => [...newItems, ...prev])
    queueMicrotask(() => newItems.forEach(kick))
  }, [woId, kick])

  // ── 4. Delete (server-side trash if uploaded, local-only otherwise)
  const deleteOne = useCallback(async (id) => {
    const item = itemsRef.current.find(i => i.id === id)
    if (!item) return
    if (item.drive_file_id) {
      const res = await fetch(
        `/api/wo-photos/${encodeURIComponent(item.drive_file_id)}`,
        { method: 'DELETE' }
      )
      const data = await res.json().catch(() => ({}))
      if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
    }
    await pendingPhotosDB.delete(id).catch(() => {})
    if (item.previewUrl) URL.revokeObjectURL(item.previewUrl)
    removeLocal(id)
  }, [removeLocal])

  // ── 5. Retry handler for error-status items ─────────────────
  const retryOne = useCallback((id) => {
    const item = itemsRef.current.find(i => i.id === id)
    if (!item) return
    patch(id, { status: 'pending', error: null })
    kick(item)
  }, [kick, patch])

  // ── 6. Resume on visibilitychange + online ──────────────────
  useEffect(() => {
    const resume = () => {
      itemsRef.current.forEach(it => {
        if (it.status === 'pending' || it.status === 'error' ||
            it.status === 'geocoding' || it.status === 'watermarking' ||
            it.status === 'uploading') {
          kick(it)
        }
      })
    }
    const onVis    = () => { if (!document.hidden) resume() }
    const onOnline = () => resume()
    document.addEventListener('visibilitychange', onVis)
    window.addEventListener('online', onOnline)
    return () => {
      document.removeEventListener('visibilitychange', onVis)
      window.removeEventListener('online', onOnline)
    }
  }, [kick])

  // Aggregated counts the form uses for submit gating.
  const uploadedCount = items.filter(i => i.status === 'uploaded').length
  const pendingCount  = items.filter(i => i.status === 'pending' ||
                                          i.status === 'geocoding' ||
                                          i.status === 'watermarking' ||
                                          i.status === 'uploading').length
  const errorCount    = items.filter(i => i.status === 'error').length

  return {
    items,
    historicLoading,
    historicError,
    uploadedCount,
    pendingCount,
    errorCount,
    addCapture,
    addLibrary,
    deleteOne,
    retryOne,
  }
}
