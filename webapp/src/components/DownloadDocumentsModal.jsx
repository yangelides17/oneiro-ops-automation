import { useEffect, useMemo, useRef, useState } from 'react'

/**
 * DownloadDocumentsModal — three-step picker that builds a zip of
 * archived documents for the prime contractor.
 *
 * Step 1: pick a mode (Unsent backlog / WO numbers / Date range).
 * Step 2: refine with contractor + doc-type + mode-specific filters.
 * Step 3: preview the matching list, choose whether to mark sent,
 *         download. After Submit a "Cancel download" button appears
 *         that aborts the in-flight zip stream via AbortController.
 *         Closing the modal does NOT abort the download — the user
 *         may want to navigate away while it finishes.
 *
 * Props:
 *   contractors — list of contractor names from the dashboard payload
 *   onClose     — called when the user explicitly closes the modal
 */

const ALL_DOC_TYPES = ['CFR', 'Production Log', 'Sign-In', 'Certified Payroll', 'Invoice']
// SI is hidden from the unsent picker — admin gets a checkbox to bundle
// SIs alongside CP instead. Other modes still expose SI as a regular pill.
const docTypesForMode = (mode) =>
  mode === 'unsent'           ? ALL_DOC_TYPES.filter(dt => dt !== 'Sign-In')
  : mode === 'payroll_period' ? ['Certified Payroll', 'Sign-In']
  : ALL_DOC_TYPES

const MODES = [
  { id: 'unsent',     title: 'Unsent backlog',
    sub: 'Every doc marked Done but not yet Sent to the prime' },
  { id: 'wo_numbers', title: 'Specific WO numbers',
    sub: 'Comma-separated list — collect everything for those WOs' },
  { id: 'date_range', title: 'Date range',
    sub: 'Every doc whose work-end date falls in the chosen window' },
  { id: 'payroll_period', title: 'Payroll period',
    sub: 'All Certified Payroll + Sign-Ins for a payroll week or month' },
]

// Recent payroll weeks (Sundays), newest first — value = ISO Sunday.
function recentWeeks(n = 16) {
  const out = []
  const base = new Date(); base.setHours(12, 0, 0, 0)
  base.setDate(base.getDate() - base.getDay())   // this week's Sunday
  const fmt = (x) => x.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })
  for (let i = 0; i < n; i++) {
    const s = new Date(base); s.setDate(base.getDate() - i * 7)
    const e = new Date(s);    e.setDate(s.getDate() + 6)
    out.push({ value: isoOf(s), label: `${fmt(s)} – ${fmt(e)}` })
  }
  return out
}
// Recent payroll months, newest first — value = YYYY-MM.
function recentMonths(n = 12) {
  const out = []
  const now = new Date()
  for (let i = 0; i < n; i++) {
    const m = new Date(now.getFullYear(), now.getMonth() - i, 1)
    out.push({
      value: `${m.getFullYear()}-${String(m.getMonth() + 1).padStart(2, '0')}`,
      label: m.toLocaleDateString('en-US', { month: 'long', year: 'numeric' }),
    })
  }
  return out
}

const isoOf = (d) => {
  const y = d.getFullYear()
  const m = String(d.getMonth() + 1).padStart(2, '0')
  const dd = String(d.getDate()).padStart(2, '0')
  return `${y}-${m}-${dd}`
}

const fmtBytes = (n) => {
  const v = Number(n) || 0
  if (v < 1024)         return v + ' B'
  if (v < 1024 * 1024)  return (v / 1024).toFixed(1) + ' KB'
  return (v / (1024 * 1024)).toFixed(1) + ' MB'
}

export default function DownloadDocumentsModal({ contractors = [], onClose }) {
  const [step,    setStep]    = useState(1)
  const [mode,    setMode]    = useState('unsent')
  // Filters always start empty across every mode — the admin
  // explicitly picks what's in the batch. Step 2 disables the
  // "Preview" button until at least one contractor and one doc type
  // is checked, so empty defaults can't accidentally fan out to a
  // huge zip.
  const [selectedContractors, setSelectedContractors] = useState([])
  const [selectedDocTypes,    setSelectedDocTypes]    = useState([])
  const [woInput,             setWoInput]             = useState('')
  const [dateStart, setDateStart] = useState(() => {
    const d = new Date(); d.setDate(d.getDate() - 6); return isoOf(d)
  })
  const [dateEnd, setDateEnd] = useState(() => isoOf(new Date()))
  // Mode-specific extras
  const [includeSIsWithCP, setIncludeSIsWithCP] = useState(false)   // unsent + CP
  const [includePhotos,    setIncludePhotos]    = useState(false)   // wo_numbers
  // payroll_period
  const [granularity, setGranularity] = useState('week')   // 'week' | 'month'
  const [periodWeek,  setPeriodWeek]  = useState('')       // ISO Sunday
  const [periodMonth, setPeriodMonth] = useState('')       // YYYY-MM
  const weekOpts  = useMemo(() => recentWeeks(16), [])
  const monthOpts = useMemo(() => recentMonths(12), [])
  // Preview
  const [previewLoading, setPreviewLoading] = useState(false)
  const [previewError,   setPreviewError]   = useState(null)
  const [preview,        setPreview]        = useState(null)
  // Submission
  const [markSent, setMarkSent] = useState(true)
  const [submitState, setSubmitState] = useState('idle')   // idle | working | cancelled | done | error
  const [submitError, setSubmitError] = useState(null)
  const abortRef = useRef(null)

  // Close on Escape (only if not mid-download)
  useEffect(() => {
    const handler = (e) => {
      if (e.key !== 'Escape') return
      onClose?.()
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [onClose])

  // Reset filters to empty whenever the mode changes — keeps the
  // selections explicit per batch instead of carrying over from a
  // prior mode the admin may have abandoned.
  useEffect(() => {
    setSelectedContractors([])
    setIncludeSIsWithCP(false)
    setIncludePhotos(false)
    if (mode === 'payroll_period') {
      // Pre-select both doc types (deselect CP for Sign-Ins only) and default
      // to the most recent week/month.
      setSelectedDocTypes(['Certified Payroll', 'Sign-In'])
      setGranularity('week')
      setPeriodWeek(recentWeeks(1)[0]?.value || '')
      setPeriodMonth(recentMonths(1)[0]?.value || '')
    } else {
      setSelectedDocTypes([])
    }
  }, [mode])

  // ── Filter payload built from current state ───────────────────
  const filters = useMemo(() => {
    const f = {
      mode,
      // wo_numbers: contractor is implicit in the WO list, skip the filter.
      contractors: mode === 'wo_numbers' ? [] : selectedContractors,
      doc_types:   selectedDocTypes,
    }
    if (mode === 'wo_numbers') {
      f.wo_ids = woInput
        .split(/[\s,]+/)
        .map(s => s.trim().toUpperCase())
        .filter(Boolean)
      if (includePhotos) f.include_photos = true
    }
    if (mode === 'date_range') {
      f.date_start = dateStart
      f.date_end   = dateEnd
    }
    if (mode === 'unsent' && selectedDocTypes.indexOf('Certified Payroll') !== -1 && includeSIsWithCP) {
      f.include_sis_with_cp = true
    }
    if (mode === 'payroll_period') {
      f.granularity = granularity
      if (granularity === 'week') f.week_start = periodWeek
      else                        f.month      = periodMonth
    }
    return f
  }, [mode, selectedContractors, selectedDocTypes, woInput, dateStart, dateEnd, includeSIsWithCP, includePhotos, granularity, periodWeek, periodMonth])

  // ── Step transitions ──────────────────────────────────────────
  const goToStep = async (next) => {
    if (next === 3) {
      // Fetch preview before showing step 3
      setPreviewLoading(true)
      setPreviewError(null)
      setPreview(null)
      try {
        const res = await fetch('/api/documents/list-batch', {
          method:  'POST',
          headers: { 'Content-Type': 'application/json' },
          body:    JSON.stringify(filters),
        })
        const data = await res.json()
        if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
        setPreview(data)
      } catch (e) {
        setPreviewError(e.message)
      } finally {
        setPreviewLoading(false)
      }
    }
    setStep(next)
  }

  // ── Step 2 validity ───────────────────────────────────────────
  const step2Valid = (() => {
    // Contractor selector is hidden for wo_numbers (contractor implicit).
    if (mode !== 'wo_numbers' && selectedContractors.length === 0) return false
    if (selectedDocTypes.length === 0)    return false
    if (mode === 'wo_numbers' && filters.wo_ids.length === 0) return false
    if (mode === 'date_range' && (!dateStart || !dateEnd))    return false
    if (mode === 'date_range' && dateStart > dateEnd)         return false
    if (mode === 'payroll_period' && granularity === 'week'  && !periodWeek)  return false
    if (mode === 'payroll_period' && granularity === 'month' && !periodMonth) return false
    return true
  })()

  // ── Submit (initiate zip download) ────────────────────────────
  const submit = async () => {
    setSubmitState('working')
    setSubmitError(null)
    const controller = new AbortController()
    abortRef.current = controller
    try {
      const res = await fetch('/api/documents/batch-download', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ ...filters, mark_sent: markSent }),
        signal:  controller.signal,
      })
      if (!res.ok) {
        const errBody = await res.json().catch(() => ({}))
        throw new Error(errBody.error || `HTTP ${res.status}`)
      }
      const blob = await res.blob()
      // Filename from Content-Disposition; fall back to a sane default
      const cd = res.headers.get('Content-Disposition') || ''
      const m  = cd.match(/filename="?([^";]+)"?/i)
      const filename = m ? m[1] : 'oneiro-docs.zip'
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = filename
      document.body.appendChild(a)
      a.click()
      a.remove()
      URL.revokeObjectURL(url)
      setSubmitState('done')
      // Auto-close after a brief success flash
      setTimeout(() => { onClose?.() }, 1200)
    } catch (e) {
      if (e.name === 'AbortError') {
        setSubmitState('cancelled')
      } else {
        setSubmitState('error')
        setSubmitError(e.message)
      }
    } finally {
      abortRef.current = null
    }
  }

  const cancelDownload = () => {
    if (abortRef.current) abortRef.current.abort()
  }

  // Toggle helpers (mutate local state)
  const toggleContractor = (c) => {
    setSelectedContractors(prev => prev.indexOf(c) === -1 ? [...prev, c] : prev.filter(x => x !== c))
  }
  const toggleDocType = (dt) => {
    setSelectedDocTypes(prev => prev.indexOf(dt) === -1 ? [...prev, dt] : prev.filter(x => x !== dt))
  }

  // ── Render ────────────────────────────────────────────────────
  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4"
      style={{ backgroundColor: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(2px)' }}
      onClick={onClose}
    >
      <div
        className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl max-h-[92vh]
                   overflow-hidden flex flex-col"
        onClick={e => e.stopPropagation()}
      >
        {/* Header */}
        <div className="px-6 pt-5 pb-3 border-b border-slate-100">
          <div className="flex items-center justify-between gap-3">
            <h2 className="text-lg font-black text-navy">📦 Download Documents</h2>
            <button
              onClick={onClose}
              className="text-slate-400 hover:text-slate-600 text-2xl leading-none"
              aria-label="Close"
            >×</button>
          </div>
          <div className="mt-2 flex gap-2 text-xs">
            {[1, 2, 3].map(n => (
              <span
                key={n}
                className={`px-2 py-0.5 rounded-full font-bold tracking-wider
                  ${step === n
                    ? 'bg-navy text-white'
                    : step > n
                      ? 'bg-green-100 text-green-700'
                      : 'bg-slate-100 text-slate-400'}`}
              >
                {step > n ? '✓ ' : ''}{n === 1 ? 'Mode' : n === 2 ? 'Filters' : 'Preview'}
              </span>
            ))}
          </div>
        </div>

        {/* Body */}
        <div className="px-6 py-5 flex-1 overflow-y-auto">
          {/* Step 1 */}
          {step === 1 && (
            <div className="space-y-3">
              <p className="text-sm text-slate-600 mb-3">
                What kind of batch are you putting together?
              </p>
              {MODES.map(m => (
                <button
                  key={m.id}
                  onClick={() => setMode(m.id)}
                  className={`w-full text-left p-4 rounded-xl border-2 transition-all
                    ${mode === m.id
                      ? 'border-navy bg-navy/5'
                      : 'border-slate-200 hover:border-navy/40'}`}
                >
                  <p className="font-bold text-slate-800">{m.title}</p>
                  <p className="text-xs text-slate-500 mt-0.5">{m.sub}</p>
                </button>
              ))}
            </div>
          )}

          {/* Step 2 */}
          {step === 2 && (
            <div className="space-y-5">
              {/* Doc types */}
              <div>
                <p className="text-xs font-extrabold uppercase tracking-widest text-slate-500 mb-2">
                  Document types
                </p>
                <div className="flex flex-wrap gap-2">
                  {(() => {
                    const opts = docTypesForMode(mode)
                    const allOn = opts.length > 0 && opts.every(dt => selectedDocTypes.indexOf(dt) !== -1)
                    return (
                      <button
                        onClick={() => setSelectedDocTypes(allOn ? [] : opts.slice())}
                        className={`text-xs px-3 py-1.5 rounded-full border font-medium transition-all
                          ${allOn
                            ? 'bg-navy text-white border-navy'
                            : 'bg-white text-slate-600 border-slate-200 hover:border-navy/40'}`}
                      >
                        {allOn ? '✓ All' : 'All'}
                      </button>
                    )
                  })()}
                  {docTypesForMode(mode).map(dt => {
                    const on = selectedDocTypes.indexOf(dt) !== -1
                    return (
                      <button
                        key={dt}
                        onClick={() => toggleDocType(dt)}
                        className={`text-xs px-3 py-1.5 rounded-full border font-medium transition-all
                          ${on
                            ? 'bg-navy text-white border-navy'
                            : 'bg-white text-slate-600 border-slate-200 hover:border-navy/40'}`}
                      >
                        {dt}
                      </button>
                    )
                  })}
                </div>

                {/* Unsent + CP → bundle SIs */}
                {mode === 'unsent' && selectedDocTypes.indexOf('Certified Payroll') !== -1 && (
                  <label className="mt-3 flex items-start gap-2 text-sm text-slate-700 cursor-pointer">
                    <input
                      type="checkbox"
                      className="mt-0.5"
                      checked={includeSIsWithCP}
                      onChange={e => setIncludeSIsWithCP(e.target.checked)}
                    />
                    <span>
                      Include matching Sign-Ins for each Certified Payroll
                      <span className="block text-[11px] text-slate-400">
                        Done Sign-Ins for the CP's contract+borough that fall within its payroll week, bundled into a sibling folder per CP.
                      </span>
                    </span>
                  </label>
                )}
              </div>

              {/* Contractors — hidden in wo_numbers mode (WO list implies contractor) */}
              {mode !== 'wo_numbers' && (
                <div>
                  <p className="text-xs font-extrabold uppercase tracking-widest text-slate-500 mb-2">
                    Contractors
                  </p>
                  {contractors.length === 0 ? (
                    <p className="text-sm text-slate-400">No contractors loaded yet — refresh the dashboard first.</p>
                  ) : (
                    <div className="flex flex-wrap gap-2">
                      {(() => {
                        const allOn = contractors.length > 0 && contractors.every(c => selectedContractors.indexOf(c) !== -1)
                        return (
                          <button
                            onClick={() => setSelectedContractors(allOn ? [] : contractors.slice())}
                            className={`text-xs px-3 py-1.5 rounded-full border font-medium transition-all
                              ${allOn
                                ? 'bg-navy text-white border-navy'
                                : 'bg-white text-slate-600 border-slate-200 hover:border-navy/40'}`}
                          >
                            {allOn ? '✓ All' : 'All'}
                          </button>
                        )
                      })()}
                      {contractors.map(c => {
                        const on = selectedContractors.indexOf(c) !== -1
                        return (
                          <button
                            key={c}
                            onClick={() => toggleContractor(c)}
                            className={`text-xs px-3 py-1.5 rounded-full border font-medium transition-all
                              ${on
                                ? 'bg-navy text-white border-navy'
                                : 'bg-white text-slate-600 border-slate-200 hover:border-navy/40'}`}
                          >
                            {c}
                          </button>
                        )
                      })}
                    </div>
                  )}
                </div>
              )}

              {/* Mode-specific */}
              {mode === 'wo_numbers' && (
                <>
                  <div>
                    <label className="block text-xs font-extrabold uppercase tracking-widest text-slate-500 mb-2">
                      WO numbers
                    </label>
                    <textarea
                      value={woInput}
                      onChange={e => setWoInput(e.target.value)}
                      placeholder="RM-43101, RM-43102, PT-12345"
                      rows={3}
                      className="field-input w-full font-mono"
                    />
                    <p className="text-[11px] text-slate-400 mt-1">
                      Comma- or whitespace-separated. Case-insensitive.
                    </p>
                  </div>
                  <label className="flex items-start gap-2 text-sm text-slate-700 cursor-pointer">
                    <input
                      type="checkbox"
                      className="mt-0.5"
                      checked={includePhotos}
                      onChange={e => setIncludePhotos(e.target.checked)}
                    />
                    <span>
                      Include site photos
                      <span className="block text-[11px] text-slate-400">
                        Bundles every image in each WO's <code className="text-[10px]">Photos/</code> folder under <code className="text-[10px]">&lt;wo&gt;/Photos/</code> in the zip. Useful for audits.
                      </span>
                    </span>
                  </label>
                </>
              )}

              {mode === 'date_range' && (
                <div className="flex gap-3 items-end flex-wrap">
                  <div>
                    <label className="block text-xs font-extrabold uppercase tracking-widest text-slate-500 mb-1">
                      Start
                    </label>
                    <input
                      type="date"
                      value={dateStart}
                      onChange={e => setDateStart(e.target.value)}
                      className="field-input"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-extrabold uppercase tracking-widest text-slate-500 mb-1">
                      End
                    </label>
                    <input
                      type="date"
                      value={dateEnd}
                      onChange={e => setDateEnd(e.target.value)}
                      className="field-input"
                    />
                  </div>
                </div>
              )}

              {mode === 'payroll_period' && (
                <div className="space-y-3">
                  <div>
                    <label className="block text-xs font-extrabold uppercase tracking-widest text-slate-500 mb-2">
                      Period
                    </label>
                    <div className="inline-flex rounded-lg border border-slate-200 overflow-hidden">
                      {['week', 'month'].map(g => (
                        <button
                          key={g}
                          type="button"
                          onClick={() => setGranularity(g)}
                          className={`text-sm font-semibold px-4 py-1.5 transition-all ${
                            granularity === g ? 'bg-navy text-white' : 'bg-white text-slate-600 hover:bg-slate-50'
                          }`}
                        >
                          {g === 'week' ? 'Payroll week' : 'Month'}
                        </button>
                      ))}
                    </div>
                  </div>
                  <div>
                    <label className="block text-xs font-extrabold uppercase tracking-widest text-slate-500 mb-1">
                      {granularity === 'week' ? 'Week' : 'Month'}
                    </label>
                    {granularity === 'week' ? (
                      <select value={periodWeek} onChange={e => setPeriodWeek(e.target.value)} className="field-input">
                        {weekOpts.map(w => <option key={w.value} value={w.value}>{w.label}</option>)}
                      </select>
                    ) : (
                      <select value={periodMonth} onChange={e => setPeriodMonth(e.target.value)} className="field-input">
                        {monthOpts.map(m => <option key={m.value} value={m.value}>{m.label}</option>)}
                      </select>
                    )}
                  </div>
                  <p className="text-[11px] text-slate-400">
                    Certified Payroll + Sign-Ins for the selected period. Deselect a doc type above to narrow it (e.g. Sign-Ins only, to reconcile a CP).
                  </p>
                </div>
              )}
            </div>
          )}

          {/* Step 3 */}
          {step === 3 && (
            <div className="space-y-4">
              {previewLoading && (
                <div className="flex items-center gap-3 text-slate-500 text-sm">
                  <div className="w-5 h-5 border-[2px] border-slate-200 border-t-navy rounded-full animate-spin" />
                  Loading preview…
                </div>
              )}
              {previewError && (
                <div className="card p-4 border-red-200 bg-red-50 text-sm text-red-700">
                  Failed to load preview: {previewError}
                </div>
              )}
              {preview && (
                <>
                  <div className="grid grid-cols-3 gap-3">
                    <div className="card p-3">
                      <p className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">
                        Files
                      </p>
                      <p className="text-2xl font-black text-navy">{preview.counts?.total ?? 0}</p>
                    </div>
                    <div className="card p-3">
                      <p className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">
                        Approx. size
                      </p>
                      <p className="text-2xl font-black text-navy">
                        {fmtBytes((preview.files || []).reduce((s, f) => s + (Number(f.size) || 0), 0))}
                      </p>
                    </div>
                    <div className="card p-3">
                      <p className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">
                        Needs review
                      </p>
                      <p className={`text-2xl font-black ${(preview.missing || []).length > 0 ? 'text-amber-600' : 'text-green-600'}`}>
                        {(preview.missing || []).length}
                      </p>
                    </div>
                  </div>

                  <div>
                    <p className="text-xs font-extrabold uppercase tracking-widest text-slate-500 mb-1.5">
                      By document type
                    </p>
                    <div className="flex flex-wrap gap-2 text-xs">
                      {Object.entries(preview.counts?.by_doc_type || {}).map(([k, v]) => (
                        <span key={k} className="bg-slate-100 text-slate-700 rounded-full px-3 py-1">
                          {k} <span className="text-slate-400">·</span> {v}
                        </span>
                      ))}
                      {Object.keys(preview.counts?.by_doc_type || {}).length === 0 && (
                        <span className="text-slate-400">No matches</span>
                      )}
                    </div>
                  </div>

                  <div>
                    <p className="text-xs font-extrabold uppercase tracking-widest text-slate-500 mb-1.5">
                      By contractor
                    </p>
                    <div className="flex flex-wrap gap-2 text-xs">
                      {Object.entries(preview.counts?.by_contractor || {}).map(([k, v]) => (
                        <span key={k} className="bg-slate-100 text-slate-700 rounded-full px-3 py-1">
                          {k} <span className="text-slate-400">·</span> {v}
                        </span>
                      ))}
                    </div>
                  </div>

                  {preview.missing && preview.missing.length > 0 && (
                    <div className="card p-3 border-amber-200 bg-amber-50">
                      <p className="text-xs font-bold text-amber-800 mb-1">
                        ⚠ {preview.missing.length} requested doc{preview.missing.length === 1 ? '' : 's'} not yet generated/approved
                      </p>
                      <div className="text-[11px] text-amber-700 max-h-32 overflow-y-auto space-y-0.5">
                        {preview.missing.slice(0, 25).map((m, i) => (
                          <p key={i} className="font-mono">
                            {(m.wo_id || [m.contract_num, m.borough, m.work_date].filter(Boolean).join(' ') || '—')} — {m.doc_type} — {m.reason}
                          </p>
                        ))}
                        {preview.missing.length > 25 && (
                          <p className="italic">…and {preview.missing.length - 25} more (full list in MANIFEST.txt)</p>
                        )}
                      </div>
                    </div>
                  )}

                  {preview.warnings && preview.warnings.length > 0 && (
                    <div className="card p-3 border-amber-200 bg-amber-50">
                      <p className="text-xs font-bold text-amber-800 mb-1">Warnings</p>
                      <div className="text-[11px] text-amber-700 space-y-0.5">
                        {preview.warnings.map((w, i) => <p key={i}>{w}</p>)}
                      </div>
                    </div>
                  )}

                  <label className="flex items-center gap-2 text-sm text-slate-700 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={markSent}
                      onChange={e => setMarkSent(e.target.checked)}
                    />
                    Mark these documents as <strong>Sent</strong> after the download completes
                  </label>

                  {submitState === 'working' && (
                    <div className="card p-3 text-sm text-slate-700 flex items-center gap-3">
                      <div className="w-5 h-5 border-[2px] border-slate-200 border-t-navy rounded-full animate-spin" />
                      Building zip… (this can take 1–2 min for large batches)
                    </div>
                  )}
                  {submitState === 'cancelled' && (
                    <div className="card p-3 border-amber-200 bg-amber-50 text-sm text-amber-800">
                      Download cancelled. No Sent flags were updated.
                    </div>
                  )}
                  {submitState === 'done' && (
                    <div className="card p-3 border-green-200 bg-green-50 text-sm text-green-800">
                      ✓ Downloaded.
                    </div>
                  )}
                  {submitState === 'error' && (
                    <div className="card p-3 border-red-200 bg-red-50 text-sm text-red-700">
                      Download failed: {submitError}
                    </div>
                  )}
                </>
              )}
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-4 border-t border-slate-100 flex items-center justify-between gap-3">
          <button
            onClick={onClose}
            className="text-sm text-slate-500 hover:text-slate-700"
          >
            Close
          </button>
          <div className="flex gap-2">
            {step > 1 && submitState !== 'working' && submitState !== 'done' && (
              <button
                onClick={() => setStep(step - 1)}
                className="btn-outline text-sm px-4 py-2"
              >
                Back
              </button>
            )}
            {step === 1 && (
              <button onClick={() => goToStep(2)} className="btn-primary text-sm px-4 py-2">
                Next
              </button>
            )}
            {step === 2 && (
              <button
                onClick={() => goToStep(3)}
                disabled={!step2Valid}
                className="btn-primary text-sm px-4 py-2 disabled:opacity-50 disabled:cursor-not-allowed"
              >
                Preview
              </button>
            )}
            {step === 3 && submitState === 'idle' && (
              <button
                onClick={submit}
                disabled={!preview || (preview.files || []).length === 0}
                className="btn-primary text-sm px-4 py-2 disabled:opacity-50 disabled:cursor-not-allowed"
              >
                Download {(preview?.counts?.total ?? 0)} file{preview?.counts?.total === 1 ? '' : 's'}
              </button>
            )}
            {step === 3 && submitState === 'working' && (
              <button
                onClick={cancelDownload}
                className="text-sm px-4 py-2 rounded-xl bg-red-500 text-white hover:bg-red-600 active:opacity-80"
              >
                Cancel download
              </button>
            )}
            {step === 3 && (submitState === 'cancelled' || submitState === 'error') && (
              <button onClick={submit} className="btn-primary text-sm px-4 py-2">
                Retry
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}
