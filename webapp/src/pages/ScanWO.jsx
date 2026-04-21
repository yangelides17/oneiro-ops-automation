import { useEffect, useRef, useState } from 'react'
import { Link } from 'react-router-dom'

// Shared HEIC → JPEG helper (same approach as FieldReport.jsx PhotoPicker).
// We convert HEIC at pick time because the Python parser on the worker
// only accepts PDF / JPEG / PNG / TIFF. HEIC previews in the browser
// need this conversion too — <img> can't decode HEIC bytes.
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

// Items: { id, name, size, type, status, file, previewUrl, error, fileId }
// Status: 'processing' | 'ready' | 'uploading' | 'uploaded' | 'error'
const newItem = (file) => ({
  id:         `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
  name:       file.name || 'scan',
  size:       file.size || 0,
  type:       file.type || '',
  status:     isHeic(file) ? 'processing' : 'ready',
  file,
  previewUrl: PREVIEWABLE_IMAGE.test(file.type || '') ? URL.createObjectURL(file) : null,
  error:      null,
  fileId:     null,
})

const formatBytes = (n) => {
  if (!n) return '0 B'
  const units = ['B', 'KB', 'MB', 'GB']
  let u = 0, v = n
  while (v >= 1024 && u < units.length - 1) { v /= 1024; u++ }
  return `${v.toFixed(v < 10 ? 1 : 0)} ${units[u]}`
}

export default function ScanWO() {
  const fileRef = useRef(null)
  // Array-of-items from day 1 so batch support later is just a UI change
  const [items,     setItems]     = useState([])
  const [submitting, setSubmitting] = useState(false)
  const [error,     setError]     = useState('')
  const [submittedCount, setSubmittedCount] = useState(0)

  // Clean up object URLs on unmount
  useEffect(() => () => {
    items.forEach(i => { if (i.previewUrl) URL.revokeObjectURL(i.previewUrl) })
  }, [])   // eslint-disable-line react-hooks/exhaustive-deps

  const addFiles = (fileList) => {
    const added = Array.from(fileList || [])
    if (added.length === 0) return

    // Filter out oversized files up front so the user sees a friendly error
    const oversized = added.filter(f => f.size > MAX_FILE_MB * 1024 * 1024)
    if (oversized.length > 0) {
      setError(`${oversized.length} file${oversized.length === 1 ? '' : 's'} over ${MAX_FILE_MB} MB skipped. Compress the PDF or split the stack.`)
    } else {
      setError('')
    }
    const accepted = added.filter(f => f.size <= MAX_FILE_MB * 1024 * 1024)
    const newItems = accepted.map(newItem)
    setItems(prev => [...prev, ...newItems])

    // Background HEIC conversion for items flagged 'processing'
    newItems.forEach(item => {
      if (item.status === 'ready') return
      convertHeicToJpeg(item.file).then(processed => {
        const converted  = processed !== item.file
        const previewUrl = URL.createObjectURL(processed)
        setItems(prev => prev.map(p => p.id === item.id
          ? { ...p, file: processed, previewUrl, type: processed.type,
              status: converted ? 'ready' : 'error',
              error:  converted ? null : 'Couldn\u2019t convert HEIC — remove and try a different format' }
          : p))
      }).catch(err => {
        setItems(prev => prev.map(p => p.id === item.id
          ? { ...p, status: 'error', error: err?.message || 'Failed to prepare file' }
          : p))
      })
    })
  }

  const onFileInput = (e) => {
    addFiles(e.target.files)
    e.target.value = ''
  }

  const [dragging, setDragging] = useState(false)
  const onDragOver = (e) => { e.preventDefault(); setDragging(true) }
  const onDragLeave = () => setDragging(false)
  const onDrop = (e) => {
    e.preventDefault()
    setDragging(false)
    addFiles(e.dataTransfer?.files)
  }

  const remove = (id) => setItems(prev => {
    const target = prev.find(p => p.id === id)
    if (target?.previewUrl) URL.revokeObjectURL(target.previewUrl)
    return prev.filter(p => p.id !== id)
  })

  const reset = () => {
    items.forEach(i => { if (i.previewUrl) URL.revokeObjectURL(i.previewUrl) })
    setItems([])
    setError('')
    setSubmittedCount(0)
  }

  // ── Upload one item. Used by initial batch submit + per-item retry
  //    + retry-all-failed. Returns 'uploaded' | 'error'.
  const uploadOne = async (item) => {
    setItems(prev => prev.map(p => p.id === item.id
      ? { ...p, status: 'uploading', error: null }
      : p))
    try {
      const form = new FormData()
      form.append('file', item.file, item.name)
      const res  = await fetch('/api/upload-wo', { method: 'POST', body: form })
      const data = await res.json().catch(() => ({}))
      if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
      setItems(prev => prev.map(p => p.id === item.id
        ? { ...p, status: 'uploaded', fileId: data.file_id }
        : p))
      return 'uploaded'
    } catch (err) {
      setItems(prev => prev.map(p => p.id === item.id
        ? { ...p, status: 'error', error: err?.message || 'Upload failed' }
        : p))
      return 'error'
    }
  }

  // Concurrency pool — N in-flight at once. Claude Vision is unaffected
  // (worker processes Scan Inbox serially); Drive writes for 20 files
  // go from ~40s sequential to ~8s parallel. Cap at 5 to stay kind to
  // mobile connections.
  const UPLOAD_CONCURRENCY = 5
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

  // Snapshot-count helper — recomputes submittedCount from the current
  // items list. setItems's functional-update form is used so we never
  // read stale state.
  const refreshSubmittedCount = () => {
    setItems(prev => {
      const ok = prev.filter(p => p.status === 'uploaded').length
      setSubmittedCount(ok)
      return prev
    })
  }

  // ── Submit all ready items ────────────────────────────────────
  const handleSubmit = async () => {
    setError('')
    if (items.length === 0) return
    if (items.some(i => i.status === 'processing')) {
      setError('Some files are still being prepared. Wait a moment and try again.')
      return
    }
    const toUpload = items.filter(i => i.status === 'ready' || i.status === 'error')
    if (toUpload.length === 0) {
      setError('Nothing to upload.')
      return
    }

    setSubmitting(true)
    await runPool(toUpload, uploadOne)
    refreshSubmittedCount()
    setSubmitting(false)
  }

  // ── Retry a single failed item ────────────────────────────────
  const retryOne = async (id) => {
    const item = items.find(i => i.id === id)
    if (!item || item.status !== 'error') return
    setSubmitting(true)
    await uploadOne(item)
    refreshSubmittedCount()
    setSubmitting(false)
  }

  // ── Retry all failed items (uses the same concurrency pool) ──
  const retryAllFailed = async () => {
    const failed = items.filter(i => i.status === 'error')
    if (failed.length === 0) return
    setError('')
    setSubmitting(true)
    await runPool(failed, uploadOne)
    refreshSubmittedCount()
    setSubmitting(false)
  }

  const processingCount = items.filter(i => i.status === 'processing').length
  const readyCount      = items.filter(i => i.status === 'ready').length
  const uploadingCount  = items.filter(i => i.status === 'uploading').length
  const uploadedCount   = items.filter(i => i.status === 'uploaded').length
  const errorCount      = items.filter(i => i.status === 'error').length
  const allUploaded     = items.length > 0 && items.every(i => i.status === 'uploaded')

  return (
    <div className="max-w-3xl mx-auto px-4 py-6 space-y-6">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-black text-navy">Scan WO</h1>
        <p className="text-sm text-slate-500 mt-1">
          Upload scanned Work Orders. Each file lands in the Scan Inbox
          and is parsed within ~1 minute. You&apos;ll see new rows appear on
          the{' '}
          <Link to="/" className="text-navy underline">Dashboard</Link>.
        </p>
      </div>

      {error && (
        <div className="bg-red-50 border border-red-200 text-red-700 text-sm
                        px-4 py-2.5 rounded-xl">
          {error}
        </div>
      )}

      {allUploaded ? (
        /* Success state */
        <div className="bg-green-50 border border-green-200 rounded-xl p-6 text-center space-y-3">
          <p className="text-3xl">✅</p>
          <p className="text-base font-bold text-green-800">
            {submittedCount} work order{submittedCount === 1 ? '' : 's'} uploaded
          </p>
          <p className="text-sm text-slate-600">
            Parsing in progress. Check the Dashboard in about a minute for
            the new row{submittedCount === 1 ? '' : 's'}.
          </p>
          <div className="flex gap-2 justify-center pt-2">
            <button onClick={reset}
              className="btn-primary text-sm px-4 py-2">
              Upload another
            </button>
            <Link to="/" className="btn-ghost text-sm px-4 py-2">
              Go to Dashboard
            </Link>
          </div>
        </div>
      ) : (
        <>
          {/* Drop zone */}
          <div
            onClick={() => !submitting && fileRef.current?.click()}
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            onDrop={onDrop}
            className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer select-none
                        transition-all ${dragging
                          ? 'border-navy bg-navy/5'
                          : 'border-slate-200 hover:border-navy/40 hover:bg-slate-50'}`}
          >
            <p className="text-3xl mb-2">📄</p>
            <p className="text-sm font-semibold text-slate-700">
              Drop files here or tap to select
            </p>
            <p className="text-xs text-slate-400 mt-1">
              PDF, JPEG, PNG, HEIC — up to {MAX_FILE_MB} MB each
            </p>
            <input
              ref={fileRef} type="file" multiple className="hidden"
              accept={ACCEPTED_MIMES} onChange={onFileInput}
            />
          </div>

          {/* Queue */}
          {items.length > 0 && (
            <div className="space-y-2">
              {items.map(item => (
                <ItemRow
                  key={item.id} item={item}
                  onRemove={remove} onRetry={retryOne}
                  disableRemove={submitting || item.status === 'uploaded'}
                  disableRetry={submitting}
                />
              ))}

              {/* Status summary */}
              <div className="flex flex-wrap gap-x-4 gap-y-1 text-[11px] font-semibold pt-1">
                {readyCount      > 0 && <span className="text-slate-600">{readyCount} ready</span>}
                {processingCount > 0 && <span className="text-slate-500">Preparing {processingCount}…</span>}
                {uploadingCount  > 0 && <span className="text-navy">Uploading {uploadingCount}…</span>}
                {uploadedCount   > 0 && <span className="text-green-700">✓ {uploadedCount} uploaded</span>}
                {errorCount      > 0 && <span className="text-red-600">✕ {errorCount} failed</span>}
              </div>

              {/* Submit + bulk retry */}
              <div className="pt-3 space-y-2">
                <button
                  onClick={handleSubmit}
                  disabled={submitting || items.length === 0 || processingCount > 0}
                  className="w-full btn-primary py-3 text-sm font-bold
                             disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {submitting
                    ? `Uploading… ${uploadedCount + errorCount} of ${items.length}`
                    : `Upload ${items.length} work order${items.length === 1 ? '' : 's'}`}
                </button>
                {errorCount > 0 && !submitting && (
                  <button
                    onClick={retryAllFailed}
                    className="w-full py-2.5 rounded-xl font-bold text-sm
                               bg-red-50 text-red-700 hover:bg-red-100 transition-colors
                               border border-red-200"
                  >
                    Retry all {errorCount} failed
                  </button>
                )}
              </div>
            </div>
          )}
        </>
      )}
    </div>
  )
}

// ── Single queue row ──────────────────────────────────────────
function ItemRow({ item, onRemove, onRetry, disableRemove, disableRetry }) {
  const isImage = (item.type || '').startsWith('image/') && item.status !== 'processing'
  const isPdf   = (item.type || '').toLowerCase().includes('pdf')

  return (
    <div className="flex items-center gap-3 border border-slate-200 rounded-xl p-2.5 bg-white">
      {/* Thumbnail */}
      <div className="w-14 h-14 rounded-lg border border-slate-200 bg-slate-50
                      flex items-center justify-center overflow-hidden flex-shrink-0">
        {item.status === 'processing' && (
          <div className="flex flex-col items-center gap-0.5">
            <div className="w-4 h-4 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
            <span className="text-[8px] text-slate-400 uppercase tracking-wider">HEIC</span>
          </div>
        )}
        {item.status === 'uploading' && (
          <div className="w-4 h-4 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
        )}
        {isImage && item.previewUrl && item.status !== 'processing' && item.status !== 'uploading' && (
          <img src={item.previewUrl} alt={item.name} className="w-full h-full object-cover" />
        )}
        {isPdf && item.status !== 'processing' && item.status !== 'uploading' && (
          <span className="text-2xl">📄</span>
        )}
      </div>

      {/* Name + status */}
      <div className="flex-1 min-w-0">
        <p className="text-sm font-semibold text-slate-800 truncate">{item.name}</p>
        <p className="text-[11px] text-slate-400 flex items-center gap-1.5">
          <span>{formatBytes(item.size)}</span>
          <span>•</span>
          <StatusLabel item={item} />
        </p>
        {item.error && (
          <p className="text-[11px] text-red-600 mt-0.5">{item.error}</p>
        )}
      </div>

      {/* Actions */}
      <div className="flex items-center gap-1 flex-shrink-0">
        {item.status === 'error' && onRetry && (
          <button
            type="button"
            onClick={() => onRetry(item.id)}
            disabled={disableRetry}
            className="text-[11px] font-bold px-2.5 py-1 rounded-lg
                       bg-red-50 text-red-700 hover:bg-red-100 transition-colors
                       border border-red-200 disabled:opacity-50 disabled:cursor-not-allowed"
            title="Retry this upload"
          >Retry</button>
        )}
        {!disableRemove && (
          <button
            type="button"
            onClick={() => onRemove(item.id)}
            className="w-7 h-7 flex items-center justify-center rounded-full
                       text-slate-400 hover:text-red-500 hover:bg-red-50 transition-colors"
            title="Remove"
          >×</button>
        )}
      </div>
    </div>
  )
}

function StatusLabel({ item }) {
  switch (item.status) {
    case 'processing': return <span className="text-slate-500">Converting HEIC…</span>
    case 'ready':      return <span className="text-slate-500">Ready to upload</span>
    case 'uploading':  return <span className="text-navy font-semibold">Uploading…</span>
    case 'uploaded':   return <span className="text-green-700 font-semibold">Uploaded — parsing</span>
    case 'error':      return <span className="text-red-600 font-semibold">Failed</span>
    default:           return null
  }
}
