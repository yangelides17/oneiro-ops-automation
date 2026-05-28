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

async function reverseGeocode(lat, lng) {
  try {
    const res = await fetch('/api/reverse-geocode', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({ lat, lng }),
    })
    const data = await res.json().catch(() => null)
    if (!data || data.error) return null
    return data
  } catch { return null }
}

async function uploadToDrive(file, wo_id) {
  const form = new FormData()
  form.append('photo', file)
  form.append('wo_id', wo_id)
  const res = await fetch('/api/upload-photo', { method: 'POST', body: form })
  const data = await res.json().catch(() => ({}))
  if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
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
      if (item.source === 'capture' && !item.watermark_applied) {
        patch(item.id, { status: 'geocoding' })
        let geoData = null
        if (item.geo && isFinite(item.geo.lat) && isFinite(item.geo.lng)) {
          geoData = await reverseGeocode(item.geo.lat, item.geo.lng)
        }
        item = fresh()
        if (!item) return
        patch(item.id, { status: 'watermarking' })
        const watermarked = await watermarkImage(item.blob, {
          timestamp: item.captured_at ? new Date(item.captured_at) : new Date(),
          addressLines: formatWatermarkAddress(geoData),
          lat: item.geo?.lat,
          lng: item.geo?.lng,
        })
        // Persist watermarked blob so a crash doesn't lose the overlay.
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
        if (item.previewUrl) URL.revokeObjectURL(item.previewUrl)
        const previewUrl = URL.createObjectURL(watermarked)
        patch(item.id, { blob: watermarked, previewUrl, watermark_applied: true })
      }

      // Phase B: upload to Drive. Compress en route.
      item = fresh()
      if (!item || !item.blob) return
      patch(item.id, { status: 'uploading' })
      const toShip = await compressForUpload(item.blob)
      const res = await uploadToDrive(toShip, woId)
      // Clear the queue row first — if we crash before the state patch
      // we'd rather lose the thumbnail than ship the photo twice.
      await pendingPhotosDB.delete(item.id).catch(() => {})
      patch(item.id, {
        status:         'uploaded',
        drive_file_id:  res.file_id,
        drive_file_url: res.file_url,
      })
    } catch (err) {
      patch(initial.id, { status: 'error', error: err?.message || 'Upload failed' })
    } finally {
      inFlightRef.current.delete(initial.id)
    }
  }, [woId, patch])

  // ── 3. Public adders ────────────────────────────────────────
  const addCapture = useCallback(async (file, geo) => {
    if (!file || !woId) return
    const id = newId()
    const captured_at = new Date().toISOString()
    await pendingPhotosDB.put({
      id, wo_id: woId, source: 'capture', blob: file,
      filename: file.name || `photo-${id}.jpg`, mime: file.type || 'image/jpeg',
      captured_at, geo: geo || null, watermark_applied: false,
    })
    const item = {
      id, source: 'capture', status: 'pending', blob: file,
      previewUrl: URL.createObjectURL(file),
      captured_at, geo: geo || null,
      filename: file.name || `photo-${id}.jpg`,
      mime: file.type || 'image/jpeg',
      watermark_applied: false,
    }
    setItems(prev => [item, ...prev])
    // Defer to next microtask so itemsRef is fresh inside kick().
    queueMicrotask(() => kick(item))
  }, [woId, kick])

  const addLibrary = useCallback(async (files) => {
    if (!woId || !files || !files.length) return
    const records = await Promise.all(Array.from(files).map(async f => {
      const id = newId()
      const rec = {
        id, wo_id: woId, source: 'library', blob: f,
        filename: f.name || `library-${id}.jpg`, mime: f.type || 'image/jpeg',
        captured_at: new Date().toISOString(),
        geo: null, watermark_applied: true,  // library imports skip watermark
      }
      await pendingPhotosDB.put(rec)
      return rec
    }))
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
