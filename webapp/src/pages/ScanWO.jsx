import { useEffect, useRef, useState } from 'react'
import { Link } from 'react-router-dom'

// ── HEIC → JPEG helper (same as FieldReport.jsx PhotoPicker). We
//    convert at pick time because the Python parser accepts only
//    PDF/JPEG/PNG/TIFF, and browsers can't preview HEIC either.
function isHeic(file) {
  const type = String(file?.type || '').toLowerCase()
  const name = String(file?.name || '').toLowerCase()
  return type === 'image/heic' || type === 'image/heif' ||
         name.endsWith('.heic') || name.endsWith('.heif')
}

async function convertHeicToJpeg(file) {
  if (!file || !isHeic(file)) return file
  try {
    const mod       = await import('heic2any')
    const heic2any  = mod.default || mod
    const converted = await heic2any({ blob: file, toType: 'image/jpeg', quality: 0.95 })
    const blob      = Array.isArray(converted) ? converted[0] : converted
    const baseName  = (file.name || 'scan').replace(/\.(heic|heif)$/i, '')
    return new File([blob], `${baseName}.jpg`, { type: 'image/jpeg' })
  } catch (err) {
    console.warn('HEIC conversion failed for WO scan:', err)
    return file
  }
}

const PREVIEWABLE_IMAGE = /^image\/(jpeg|png)$/i
const MAX_FILE_MB       = 15
const ACCEPTED_MIMES    = 'application/pdf,image/jpeg,image/png,image/heic,image/heif,image/tiff'
const POLL_INTERVAL_MS  = 10_000

const formatBytes = (n) => {
  if (!n) return '0 B'
  const units = ['B', 'KB', 'MB', 'GB']
  let u = 0, v = n
  while (v >= 1024 && u < units.length - 1) { v /= 1024; u++ }
  return `${v.toFixed(v < 10 ? 1 : 0)} ${units[u]}`
}

const todayIso = () => {
  const d = new Date()
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
}

const historyKey = (iso = todayIso()) => `scanwo-history-${iso}`

// Purge any localStorage keys from previous days so the admin's
// history doesn't accumulate indefinitely.
function purgeOldHistory() {
  const today = todayIso()
  try {
    for (const key of Object.keys(localStorage)) {
      if (!key.startsWith('scanwo-history-')) continue
      if (key === historyKey(today)) continue
      localStorage.removeItem(key)
    }
  } catch { /* non-critical */ }
}

// Load today's history; repair any zombie in-progress items that
// survived a page reload (they can't actually finish because their
// File object is gone).
function loadTodayHistory() {
  try {
    const raw = localStorage.getItem(historyKey())
    if (!raw) return []
    const parsed = JSON.parse(raw)
    if (!Array.isArray(parsed)) return []
    return parsed.map(item => {
      if (item.parseStatus === 'done') return item
      if (item.uploadStatus === 'uploading' || item.parseStatus === 'parsing') {
        return {
          ...item,
          uploadStatus: 'upload_error',
          parseStatus:  undefined,
          error:        'Page reload interrupted this upload — delete and re-pick the file.',
        }
      }
      return item
    })
  } catch {
    return []
  }
}

function persistHistory(items) {
  // Only keep serializable fields; File objects never survive JSON anyway
  const clean = items.map(({ file, ...rest }) => rest)   // eslint-disable-line no-unused-vars
  try {
    localStorage.setItem(historyKey(), JSON.stringify(clean))
  } catch { /* quota / private mode */ }
}

// Ready-queue item shape:
//   { id, name, size, type, status: 'processing'|'ready'|'error', file, previewUrl, error }
function newReadyItem(file) {
  return {
    id:         `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    name:       file.name || 'scan',
    size:       file.size || 0,
    type:       file.type || '',
    status:     isHeic(file) ? 'processing' : 'ready',
    file,
    previewUrl: PREVIEWABLE_IMAGE.test(file.type || '') ? URL.createObjectURL(file) : null,
    error:      null,
  }
}

// History-queue item shape:
//   { id, name, size, type, uploadedAt, fileId,
//     uploadStatus: 'uploading' | 'uploaded' | 'upload_error',
//     parseStatus:  'parsing' | 'done' | 'error' | undefined,
//     woIds: string[], error, expanded, file (runtime only) }

export default function ScanWO() {
  const fileRef    = useRef(null)
  const [readyItems,   setReadyItems]   = useState([])
  const [historyItems, setHistoryItems] = useState(loadTodayHistory)
  const [submitting,   setSubmitting]   = useState(false)
  const [error,        setError]        = useState('')
  const [dragging,     setDragging]     = useState(false)

  // Clean up stale day keys on mount
  useEffect(() => { purgeOldHistory() }, [])

  // One-shot catch-up poll on mount. If the user refreshed while
  // something was in flight (or just popped back in an hour later),
  // the persisted items may be stuck in 'upload_error' / 'parsing'.
  // Ask the server for the current state of every item that still
  // has a fileId and isn't confirmed done.
  useEffect(() => {
    const candidates = loadTodayHistory().filter(h => h.fileId && h.parseStatus !== 'done')
    if (candidates.length === 0) return
    let cancelled = false
    ;(async () => {
      try {
        const res = await fetch('/api/scan-status', {
          method:  'POST',
          headers: { 'Content-Type': 'application/json' },
          body:    JSON.stringify({ file_ids: candidates.map(c => c.fileId) }),
        })
        const data = await res.json().catch(() => ({}))
        if (cancelled || !Array.isArray(data.statuses)) return
        setHistoryItems(prev => prev.map(p => {
          const s = data.statuses.find(x => x.file_id === p.fileId)
          if (!s) return p
          if (s.status === 'done') {
            return { ...p, parseStatus: 'done', uploadStatus: 'uploaded',
                     error: null, woIds: s.wo_ids || [] }
          }
          if (s.status === 'error') {
            return { ...p, parseStatus: 'error', error: s.message || 'Parse failed' }
          }
          return p
        }))
      } catch (err) {
        console.warn('scan-status catch-up poll failed:', err)
      }
    })()
    return () => { cancelled = true }
  }, [])   // eslint-disable-line react-hooks/exhaustive-deps

  // Persist history whenever it changes
  useEffect(() => { persistHistory(historyItems) }, [historyItems])

  // Clean up object URLs on unmount
  useEffect(() => () => {
    readyItems.forEach(i => { if (i.previewUrl) URL.revokeObjectURL(i.previewUrl) })
  }, [])   // eslint-disable-line react-hooks/exhaustive-deps

  // ── Ready queue: add / remove ─────────────────────────────────
  const addFiles = (fileList) => {
    const added = Array.from(fileList || [])
    if (added.length === 0) return

    const oversized = added.filter(f => f.size > MAX_FILE_MB * 1024 * 1024)
    if (oversized.length > 0) {
      setError(`${oversized.length} file${oversized.length === 1 ? '' : 's'} over ${MAX_FILE_MB} MB skipped. Compress the PDF or split the stack.`)
    } else {
      setError('')
    }
    const accepted = added.filter(f => f.size <= MAX_FILE_MB * 1024 * 1024)
    const next = accepted.map(newReadyItem)
    setReadyItems(prev => [...prev, ...next])

    // Background HEIC conversion for items flagged 'processing'
    next.forEach(item => {
      if (item.status === 'ready') return
      convertHeicToJpeg(item.file).then(processed => {
        const converted  = processed !== item.file
        const previewUrl = URL.createObjectURL(processed)
        setReadyItems(prev => prev.map(p => p.id === item.id
          ? { ...p, file: processed, previewUrl, type: processed.type,
              status: converted ? 'ready' : 'error',
              error:  converted ? null : 'Couldn\u2019t convert HEIC — remove and try a different format' }
          : p))
      }).catch(err => {
        setReadyItems(prev => prev.map(p => p.id === item.id
          ? { ...p, status: 'error', error: err?.message || 'Failed to prepare file' }
          : p))
      })
    })
  }

  const onFileInput = (e) => { addFiles(e.target.files); e.target.value = '' }
  const onDragOver  = (e) => { e.preventDefault(); setDragging(true) }
  const onDragLeave = () => setDragging(false)
  const onDrop      = (e) => { e.preventDefault(); setDragging(false); addFiles(e.dataTransfer?.files) }

  const removeReady = (id) => setReadyItems(prev => {
    const target = prev.find(p => p.id === id)
    if (target?.previewUrl) URL.revokeObjectURL(target.previewUrl)
    return prev.filter(p => p.id !== id)
  })

  const removeHistory = (id) => setHistoryItems(prev => prev.filter(p => p.id !== id))

  const toggleExpand = (id) => setHistoryItems(prev => prev.map(p =>
    p.id === id ? { ...p, expanded: !p.expanded } : p))

  // ── Single upload: Ready item → history item ──────────────────
  const UPLOAD_CONCURRENCY = 5

  const uploadOne = async (readyItem) => {
    // Create the history item immediately so the user sees it move
    const historyId = `${readyItem.id}-h`
    const baseHistory = {
      id:           historyId,
      readyId:      readyItem.id,
      name:         readyItem.name,
      size:         readyItem.size,
      type:         readyItem.type,
      uploadedAt:   new Date().toISOString(),
      fileId:       null,
      uploadStatus: 'uploading',
      parseStatus:  undefined,
      woIds:        [],
      error:        null,
      expanded:     false,
      file:         readyItem.file,   // kept in memory only for retry
    }
    setHistoryItems(prev => [baseHistory, ...prev])

    try {
      const form = new FormData()
      form.append('file', readyItem.file, readyItem.name)
      const res  = await fetch('/api/upload-wo', { method: 'POST', body: form })
      const data = await res.json().catch(() => ({}))
      if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
      setHistoryItems(prev => prev.map(p => p.id === historyId
        ? { ...p, uploadStatus: 'uploaded', parseStatus: 'parsing', fileId: data.file_id }
        : p))
      return 'uploaded'
    } catch (err) {
      setHistoryItems(prev => prev.map(p => p.id === historyId
        ? { ...p, uploadStatus: 'upload_error', error: err?.message || 'Upload failed' }
        : p))
      return 'error'
    }
  }

  const runPool = async (queue, worker) => {
    let cursor = 0
    const take = async () => {
      while (cursor < queue.length) {
        const idx = cursor++
        await worker(queue[idx])
      }
    }
    const n = Math.min(UPLOAD_CONCURRENCY, queue.length)
    await Promise.all(Array.from({ length: n }, take))
  }

  // ── Submit the Ready queue ────────────────────────────────────
  const handleSubmit = async () => {
    setError('')
    if (readyItems.length === 0) return
    if (readyItems.some(i => i.status === 'processing')) {
      setError('Some files are still being prepared. Wait a moment and try again.')
      return
    }
    const toUpload = readyItems.filter(i => i.status === 'ready' || i.status === 'error')
    if (toUpload.length === 0) {
      setError('Nothing to upload.')
      return
    }

    setSubmitting(true)
    await runPool(toUpload, uploadOne)
    // Clear the Ready queue (successful + errored ones moved to history)
    setReadyItems(prev => {
      prev.forEach(i => { if (i.previewUrl) URL.revokeObjectURL(i.previewUrl) })
      return []
    })
    setSubmitting(false)
  }

  // ── Status polling ────────────────────────────────────────────
  useEffect(() => {
    const parsing = historyItems.filter(h => h.parseStatus === 'parsing' && h.fileId)
    if (parsing.length === 0) return undefined

    let cancelled = false
    const tick = async () => {
      try {
        const res  = await fetch('/api/scan-status', {
          method:  'POST',
          headers: { 'Content-Type': 'application/json' },
          body:    JSON.stringify({ file_ids: parsing.map(p => p.fileId) }),
        })
        const data = await res.json().catch(() => ({}))
        if (cancelled || !Array.isArray(data.statuses)) return
        setHistoryItems(prev => prev.map(p => {
          if (p.parseStatus !== 'parsing') return p
          const s = data.statuses.find(x => x.file_id === p.fileId)
          if (!s) return p
          if (s.status === 'done') {
            return { ...p, parseStatus: 'done', woIds: s.wo_ids || [] }
          }
          if (s.status === 'error') {
            return { ...p, parseStatus: 'error', error: s.message || 'Parse failed' }
          }
          // pending or unknown — keep polling
          return p
        }))
      } catch (err) {
        // Transient network — try again on next tick
        console.warn('scan-status poll failed:', err)
      }
    }

    // Poll immediately so the user doesn't wait a full interval
    tick()
    const interval = setInterval(tick, POLL_INTERVAL_MS)
    return () => { cancelled = true; clearInterval(interval) }
  }, [historyItems])

  // ── Summary counter ───────────────────────────────────────────
  const woCountToday = historyItems.reduce((acc, h) =>
    h.parseStatus === 'done' ? acc + (h.woIds?.length || 0) : acc, 0)

  const readyProcessing = readyItems.filter(i => i.status === 'processing').length

  return (
    <div className="max-w-3xl mx-auto px-4 py-6 space-y-6">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-black text-navy">Scan WO</h1>
        <p className="text-sm text-slate-500 mt-1">
          Upload scanned Work Orders (single or combined stacks). Each
          file is parsed within ~1 minute. See new rows on the{' '}
          <Link to="/" className="text-navy underline">Dashboard</Link>.
        </p>
      </div>

      {error && (
        <div className="bg-red-50 border border-red-200 text-red-700 text-sm px-4 py-2.5 rounded-xl">
          {error}
        </div>
      )}

      {/* Drop zone */}
      <div
        onClick={() => !submitting && fileRef.current?.click()}
        onDragOver={onDragOver} onDragLeave={onDragLeave} onDrop={onDrop}
        className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer select-none transition-all
                    ${dragging ? 'border-navy bg-navy/5' : 'border-slate-200 hover:border-navy/40 hover:bg-slate-50'}`}
      >
        <p className="text-3xl mb-2">📄</p>
        <p className="text-sm font-semibold text-slate-700">Drop files here or tap to select</p>
        <p className="text-xs text-slate-400 mt-1">PDF, JPEG, PNG, HEIC — up to {MAX_FILE_MB} MB each</p>
        <input ref={fileRef} type="file" multiple className="hidden"
               accept={ACCEPTED_MIMES} onChange={onFileInput} />
      </div>

      {/* Ready-to-submit queue */}
      {readyItems.length > 0 && (
        <section className="space-y-2">
          <h2 className="text-[11px] font-extrabold uppercase tracking-wider text-slate-500">
            Ready to Submit ({readyItems.length})
          </h2>
          {readyItems.map(item => (
            <ReadyRow key={item.id} item={item} onRemove={removeReady}
                      disableRemove={submitting} />
          ))}
          <div className="pt-2">
            <button
              onClick={handleSubmit}
              disabled={submitting || readyItems.length === 0 || readyProcessing > 0}
              className="w-full btn-primary py-3 text-sm font-bold disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {submitting
                ? 'Uploading…'
                : `Upload ${readyItems.length} work order${readyItems.length === 1 ? '' : 's'}`}
            </button>
          </div>
        </section>
      )}

      {/* Today's uploads */}
      {historyItems.length > 0 && (
        <section className="space-y-2">
          <div className="flex items-baseline justify-between">
            <h2 className="text-[11px] font-extrabold uppercase tracking-wider text-slate-500">
              Today&apos;s Uploads
            </h2>
            <span className="text-xs font-bold text-green-700">
              {woCountToday} WO{woCountToday === 1 ? '' : 's'} added today
            </span>
          </div>
          {historyItems.map(item => (
            <HistoryRow key={item.id} item={item}
                        onRemove={removeHistory}
                        onToggle={toggleExpand} />
          ))}
        </section>
      )}
    </div>
  )
}

// ── Ready-queue row ───────────────────────────────────────────
function ReadyRow({ item, onRemove, disableRemove }) {
  const isImage = (item.type || '').startsWith('image/') && item.status !== 'processing'
  const isPdf   = (item.type || '').toLowerCase().includes('pdf')
  return (
    <div className="flex items-center gap-3 border border-slate-200 rounded-xl p-2.5 bg-white">
      <div className="w-14 h-14 rounded-lg border border-slate-200 bg-slate-50
                      flex items-center justify-center overflow-hidden flex-shrink-0">
        {item.status === 'processing' && (
          <div className="flex flex-col items-center gap-0.5">
            <div className="w-4 h-4 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
            <span className="text-[8px] text-slate-400 uppercase tracking-wider">HEIC</span>
          </div>
        )}
        {isImage && item.previewUrl && item.status !== 'processing' && (
          <img src={item.previewUrl} alt={item.name} className="w-full h-full object-cover" />
        )}
        {isPdf && item.status !== 'processing' && (
          <span className="text-2xl">📄</span>
        )}
      </div>
      <div className="flex-1 min-w-0">
        <p className="text-sm font-semibold text-slate-800 truncate">{item.name}</p>
        <p className="text-[11px] text-slate-400 flex items-center gap-1.5">
          <span>{formatBytes(item.size)}</span>
          <span>•</span>
          <span>
            {item.status === 'processing' && 'Converting HEIC…'}
            {item.status === 'ready'      && 'Ready to upload'}
            {item.status === 'error'      && <span className="text-red-600 font-semibold">{item.error || 'Error'}</span>}
          </span>
        </p>
      </div>
      {!disableRemove && (
        <button type="button" onClick={() => onRemove(item.id)}
          className="w-7 h-7 flex items-center justify-center rounded-full
                     text-slate-400 hover:text-red-500 hover:bg-red-50 transition-colors flex-shrink-0"
          title="Remove">×</button>
      )}
    </div>
  )
}

// ── History row (uploaded / parsing / done / error) ───────────
function HistoryRow({ item, onRemove, onToggle }) {
  const isError   = item.uploadStatus === 'upload_error' || item.parseStatus === 'error'
  const isPending = item.uploadStatus === 'uploading' || item.parseStatus === 'parsing'
  const isDone    = item.parseStatus === 'done'
  const multiWo   = isDone && (item.woIds?.length || 0) > 1

  return (
    <div className={`border rounded-xl bg-white transition-all
                     ${isError   ? 'border-red-200 bg-red-50/30' :
                       isDone    ? 'border-green-200 bg-green-50/20' :
                                   'border-slate-200'}`}>
      <div className="flex items-start gap-3 p-2.5">
        <div className="w-9 h-9 rounded-lg bg-slate-50 border border-slate-200
                        flex items-center justify-center flex-shrink-0 mt-0.5">
          {isPending && (
            <div className="w-4 h-4 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
          )}
          {isDone && <span className="text-green-700 text-lg">✓</span>}
          {isError && <span className="text-red-600 text-lg">✕</span>}
        </div>

        <div className="flex-1 min-w-0 space-y-0.5">
          <p className="text-sm font-semibold text-slate-800 truncate">{item.name}</p>
          <div className="text-[11px] flex flex-wrap items-center gap-x-1.5 gap-y-0.5">
            {item.uploadStatus === 'uploading' && <span className="text-slate-500">Uploading…</span>}
            {item.uploadStatus === 'uploaded' && item.parseStatus === 'parsing' && <span className="text-navy font-semibold">Parsing…</span>}
            {item.uploadStatus === 'upload_error' && (
              <span className="text-red-600 font-semibold">Upload failed — {item.error}</span>
            )}
            {item.parseStatus === 'error' && (
              <span className="text-red-600 font-semibold">Parse failed — {item.error}</span>
            )}
            {isDone && !multiWo && (
              <span className="text-green-700 font-semibold">{item.woIds[0]}</span>
            )}
            {isDone && multiWo && (
              <button type="button" onClick={() => onToggle(item.id)}
                className="text-green-700 font-semibold flex items-center gap-1 hover:underline">
                {item.woIds.length} WOs <span>{item.expanded ? '⌃' : '⌄'}</span>
              </button>
            )}
          </div>
          {isError && (
            <p className="text-[11px] text-slate-500 leading-snug">
              Dismiss this entry and re-upload the file.
            </p>
          )}
        </div>

        {isError && (
          <button type="button" onClick={() => onRemove(item.id)}
            className="w-7 h-7 flex items-center justify-center rounded-full flex-shrink-0
                       text-slate-400 hover:text-red-500 hover:bg-red-50 transition-colors"
            title="Dismiss">×</button>
        )}
      </div>

      {/* Multi-WO expansion — show each WO# one per line */}
      {isDone && multiWo && item.expanded && (
        <ul className="px-3 pb-3 pt-0 space-y-0.5">
          {item.woIds.map(wo => (
            <li key={wo} className="text-xs text-slate-700 font-mono pl-12">
              └ {wo}
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}
