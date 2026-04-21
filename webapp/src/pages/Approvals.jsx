import { useEffect, useMemo, useRef, useState } from 'react'
import { Link } from 'react-router-dom'
import { Document, Page, pdfjs } from 'react-pdf'
import 'react-pdf/dist/Page/AnnotationLayer.css'
import 'react-pdf/dist/Page/TextLayer.css'

// Wire the pdf.js worker. Using the CDN URL with the same version
// react-pdf ships with keeps bundle size down and avoids Vite worker
// resolution quirks. If the bundler inlines a different pdfjs version,
// the version check at load time will warn in the console.
pdfjs.GlobalWorkerOptions.workerSrc =
  `https://unpkg.com/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`

// ── Doc-type config — filter values + display labels/colors ──
const DOC_TYPES = [
  { value: 'all',               label: 'All',             badge: '',            chip: 'bg-slate-200 text-slate-700' },
  { value: 'signin',            label: 'Sign-In',         badge: 'Sign-In',     chip: 'bg-sky-100 text-sky-800' },
  { value: 'production_log',    label: 'Production Log',  badge: 'Prod Log',    chip: 'bg-amber-100 text-amber-800' },
  { value: 'field_report',      label: 'CFR',             badge: 'CFR',         chip: 'bg-violet-100 text-violet-800' },
  { value: 'certified_payroll', label: 'Cert Payroll',    badge: 'Cert Payroll',chip: 'bg-emerald-100 text-emerald-800' },
]
const docTypeMeta = (t) => DOC_TYPES.find(x => x.value === t) || DOC_TYPES[0]

const fmtTime = (iso) => {
  if (!iso) return ''
  try {
    const d = new Date(iso)
    return d.toLocaleString(undefined, {
      month: 'numeric', day: 'numeric',
      hour: 'numeric',  minute: '2-digit',
    })
  } catch { return '' }
}

export default function Approvals() {
  const [approvals, setApprovals] = useState(null)   // null = initial loading; [] = loaded
  const [loadError, setLoadError] = useState('')
  const [filter,    setFilter]    = useState('all')
  const [selectedId, setSelectedId] = useState(null)
  const [approving,  setApproving]  = useState(false)
  const [actionError, setActionError] = useState('')
  const containerRef = useRef(null)
  const [viewerWidth, setViewerWidth] = useState(700)

  // Fetch pending approvals once on mount + after each approve action
  const refresh = async () => {
    try {
      const res  = await fetch('/api/approvals')
      const data = await res.json()
      if (data.error) throw new Error(data.error)
      setApprovals(Array.isArray(data.approvals) ? data.approvals : [])
      setLoadError('')
    } catch (err) {
      setLoadError(err.message || 'Failed to load approvals')
      setApprovals([])
    }
  }
  useEffect(() => { refresh() }, [])

  // Keep the viewer width in sync with its container (react-pdf needs
  // an explicit width to scale pages; letting it auto-fill the flex
  // column requires a measured number).
  useEffect(() => {
    const el = containerRef.current
    if (!el) return
    const ro = new ResizeObserver(entries => {
      const w = entries[0]?.contentRect?.width || 700
      setViewerWidth(Math.max(300, Math.floor(w) - 32))   // 32 = inner padding
    })
    ro.observe(el)
    return () => ro.disconnect()
  }, [])

  // Filtered list derived from approvals + filter
  const filtered = useMemo(() => {
    if (!approvals) return []
    if (filter === 'all') return approvals
    return approvals.filter(a => a.doc_type === filter)
  }, [approvals, filter])

  // Auto-select first filtered item when list changes + no selection
  useEffect(() => {
    if (!selectedId && filtered.length > 0) {
      setSelectedId(filtered[0].file_id)
    }
    // If selection isn't in filtered view (filter changed), move to first
    if (selectedId && filtered.length > 0 &&
        !filtered.some(a => a.file_id === selectedId)) {
      setSelectedId(filtered[0].file_id)
    }
    // If list is empty, clear selection
    if (filtered.length === 0 && selectedId) {
      setSelectedId(null)
    }
  }, [filtered, selectedId])

  const selected = approvals?.find(a => a.file_id === selectedId) || null

  // Approve: move file, remove locally, auto-advance to next in filtered view
  const handleApprove = async () => {
    if (!selected || approving) return
    setActionError('')
    setApproving(true)
    // Compute next-selected from CURRENT filtered list before mutation
    const idx = filtered.findIndex(a => a.file_id === selected.file_id)
    const nextItem = filtered[idx + 1] || filtered[idx - 1] || null
    try {
      const res = await fetch(
        `/api/approvals/${encodeURIComponent(selected.file_id)}/approve`,
        { method: 'POST' })
      const data = await res.json()
      if (data.error) throw new Error(data.error)
      // Remove from list; auto-advance
      setApprovals(prev => (prev || []).filter(a => a.file_id !== selected.file_id))
      setSelectedId(nextItem?.file_id || null)
    } catch (err) {
      setActionError(err.message || 'Approval failed')
    } finally {
      setApproving(false)
    }
  }

  const openInDrive = () => {
    if (!selected) return
    window.open(`https://drive.google.com/file/d/${selected.file_id}/view`, '_blank')
  }

  const pendingCount = approvals?.length ?? 0

  return (
    <div className="max-w-6xl mx-auto px-4 py-6 space-y-4">
      {/* Header */}
      <div className="flex items-baseline justify-between">
        <div>
          <h1 className="text-2xl font-black text-navy">Approvals</h1>
          <p className="text-sm text-slate-500 mt-1">
            Review and approve generated documents. Approving moves the
            file to <span className="font-mono">✅ Approved Docs</span>, which
            triggers the email + archive job within ~10 minutes.
          </p>
        </div>
        <span className="text-xs font-bold text-slate-600">
          {pendingCount} pending
        </span>
      </div>

      {/* Filter pills */}
      <div className="flex flex-wrap gap-2">
        {DOC_TYPES.map(dt => {
          const n = dt.value === 'all'
            ? pendingCount
            : (approvals || []).filter(a => a.doc_type === dt.value).length
          const active = filter === dt.value
          return (
            <button
              key={dt.value}
              onClick={() => setFilter(dt.value)}
              className={`px-3 py-1.5 rounded-full text-xs font-bold transition-colors
                ${active ? 'bg-navy text-white' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}`}
            >
              {dt.label} {n > 0 && <span className={active ? 'opacity-80' : 'opacity-70'}>· {n}</span>}
            </button>
          )
        })}
      </div>

      {loadError && (
        <div className="bg-red-50 border border-red-200 text-red-700 text-sm px-4 py-2.5 rounded-xl">
          {loadError}
        </div>
      )}

      {/* Two-pane layout */}
      <div className="grid grid-cols-1 md:grid-cols-[320px_1fr] gap-4">
        {/* List */}
        <section className="bg-white rounded-xl border border-slate-200 overflow-hidden">
          {approvals === null && (
            <p className="p-4 text-sm text-slate-500">Loading…</p>
          )}
          {approvals !== null && filtered.length === 0 && (
            <div className="p-6 text-center space-y-2">
              <p className="text-3xl">✅</p>
              <p className="text-sm font-semibold text-slate-700">All caught up</p>
              <p className="text-xs text-slate-500">
                Nothing pending in {filter === 'all' ? 'any category' : docTypeMeta(filter).label}.
              </p>
              <Link to="/" className="inline-block pt-2 text-xs text-navy underline">
                Back to Dashboard
              </Link>
            </div>
          )}
          {filtered.map(item => (
            <ApprovalRow
              key={item.file_id}
              item={item}
              active={item.file_id === selectedId}
              onClick={() => setSelectedId(item.file_id)}
            />
          ))}
        </section>

        {/* Viewer */}
        <section
          ref={containerRef}
          className="bg-white rounded-xl border border-slate-200 p-4 flex flex-col min-h-[60vh]"
        >
          {!selected && (
            <div className="flex-1 flex items-center justify-center text-sm text-slate-400">
              {filtered.length === 0
                ? 'All caught up for this filter.'
                : 'Select a document to review.'}
            </div>
          )}

          {selected && (
            <>
              <div className="flex items-start justify-between gap-3 mb-3 flex-wrap">
                <div>
                  <p className="text-sm font-bold text-slate-800 truncate">
                    {selected.filename}
                  </p>
                  <p className="text-[11px] text-slate-500 mt-0.5">
                    {docTypeMeta(selected.doc_type).label}
                    {selected.subtitle && <> · {selected.subtitle}</>}
                    <> · {fmtTime(selected.created_at)}</>
                  </p>
                </div>
                <div className="flex items-center gap-2 flex-shrink-0">
                  <button
                    type="button"
                    onClick={openInDrive}
                    className="text-xs font-bold px-3 py-1.5 rounded-lg
                               bg-slate-100 text-slate-700 hover:bg-slate-200 transition-colors"
                  >
                    Open in Drive
                  </button>
                  <button
                    type="button"
                    onClick={handleApprove}
                    disabled={approving}
                    className="text-xs font-bold px-4 py-1.5 rounded-lg
                               bg-green-600 text-white hover:bg-green-700 transition-colors
                               disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    {approving ? 'Approving…' : 'Approve'}
                  </button>
                </div>
              </div>

              {actionError && (
                <div className="bg-red-50 border border-red-200 text-red-700 text-xs
                                px-3 py-2 rounded-lg mb-3">
                  {actionError}
                </div>
              )}

              <PDFViewer fileId={selected.file_id} width={viewerWidth} />
            </>
          )}
        </section>
      </div>
    </div>
  )
}

// ── List row ──────────────────────────────────────────────────
function ApprovalRow({ item, active, onClick }) {
  const meta = docTypeMeta(item.doc_type)
  return (
    <button
      type="button"
      onClick={onClick}
      className={`w-full text-left border-b border-slate-100 px-3 py-3 transition-colors
                  ${active ? 'bg-navy/5' : 'hover:bg-slate-50'}`}
    >
      <div className="flex items-start justify-between gap-2 mb-1">
        <span className={`text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded-full ${meta.chip}`}>
          {meta.badge}
        </span>
        <span className="text-[10px] text-slate-400 flex-shrink-0">{fmtTime(item.created_at)}</span>
      </div>
      <p className="text-sm font-semibold text-slate-800 truncate">
        {item.subtitle || item.filename}
      </p>
      {item.subtitle && item.filename && item.subtitle !== item.filename.replace(/\.pdf$/i, '') && (
        <p className="text-[11px] text-slate-400 truncate">{item.filename}</p>
      )}
    </button>
  )
}

// ── PDF viewer wrapping react-pdf ─────────────────────────────
function PDFViewer({ fileId, width }) {
  const [numPages, setNumPages] = useState(null)
  const [loadErr, setLoadErr]   = useState(null)

  // react-pdf's `file` prop. Using the URL directly lets pdf.js stream
  // the document — faster than fetching the whole blob first.
  const file = useMemo(() => ({
    url: `/api/approvals/${encodeURIComponent(fileId)}/pdf`,
  }), [fileId])

  // Reset error / page count whenever fileId changes
  useEffect(() => {
    setLoadErr(null)
    setNumPages(null)
  }, [fileId])

  return (
    <div className="flex-1 min-h-0 overflow-auto bg-slate-50 rounded-lg border border-slate-200">
      {loadErr && (
        <div className="p-6 text-sm text-red-700">
          Couldn&apos;t render this PDF: {loadErr.message || String(loadErr)}
        </div>
      )}
      {!loadErr && (
        <Document
          key={fileId}
          file={file}
          onLoadSuccess={({ numPages }) => setNumPages(numPages)}
          onLoadError={(e) => setLoadErr(e)}
          loading={<div className="p-6 text-sm text-slate-500">Loading document…</div>}
          className="flex flex-col items-center gap-2 p-2"
        >
          {Array.from({ length: numPages || 0 }, (_, i) => (
            <Page
              key={i + 1}
              pageNumber={i + 1}
              width={width}
              renderTextLayer={false}
              renderAnnotationLayer={true}
              className="shadow-sm border border-slate-200 bg-white"
            />
          ))}
        </Document>
      )}
    </div>
  )
}
