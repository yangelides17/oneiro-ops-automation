import { useEffect, useMemo, useState } from 'react'
import { Document, Page } from 'react-pdf'

// Upload signed month-end paperwork (Employee Utilization / Certificates).
//
// These forms leave the building to be wet-signed, so they come back as a
// scan — often one stack covering several contracts. This modal takes that
// stack, asks the server to propose a page→document split, and shows the
// proposal for confirmation before anything is filed.
//
// The confirmation step is the point. A compliance certificate archived
// under the wrong contract is worse than one left unfiled, so nothing is
// ever committed on the model's say-so: a wrong guess costs one dropdown
// click. A single-document PDF simply comes back as one group, so dropping
// one form and dropping a whole stack are the same flow.
//
// Props:
//   candidates  — month_end_awaiting rows (used for the empty-state copy)
//   onClose()   — dismiss
//   onUploaded(uploaded) — fired after a successful commit

const MAX_BYTES = 20 * 1024 * 1024   // matches the server multer limit

const UNASSIGNED = ''

const docLabel = (c) =>
  `${c.label} · ${c.contract_num}-${c.borough}${c.contractor ? ` · ${c.contractor}` : ''}`

export default function MonthEndUploadModal({ candidates = [], onClose, onUploaded }) {
  const [step, setStep]     = useState('choose')  // choose | analyzing | review | committing
  const [error, setError]   = useState('')
  const [file, setFile]     = useState(null)
  const [preview, setPreview] = useState(null)    // server split-preview payload
  const [assign, setAssign] = useState({})        // { [pageNumber]: doc_id | '' }

  const busy = step === 'analyzing' || step === 'committing'

  useEffect(() => {
    const handler = (e) => { if (e.key === 'Escape' && !busy) onClose?.() }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [onClose, busy])

  // Candidate list for the dropdowns. Prefer the set the server echoed
  // back with the proposal — it's the same list the model chose from, so
  // the options can never drift from what was actually considered.
  const options = preview?.candidates?.length ? preview.candidates : candidates

  const analyze = async (f) => {
    if (!f) return
    if (f.size > MAX_BYTES) {
      setError(`File is too large (${(f.size / 1e6).toFixed(1)} MB). Max is 20 MB — split the stack and upload it in parts.`)
      return
    }
    setFile(f); setError(''); setStep('analyzing')
    try {
      const fd = new FormData()
      fd.append('file', f)
      const res = await fetch('/api/approvals/month-end/split-preview', { method: 'POST', body: fd })
      const json = await res.json()
      if (!res.ok || json.error) throw new Error(json.error || `HTTP ${res.status}`)

      // Seed the editable map from the proposal. Pages the server couldn't
      // place start unassigned rather than guessed.
      const seeded = {}
      for (let p = 1; p <= json.page_count; p++) seeded[p] = UNASSIGNED
      ;(json.groups || []).forEach(g => g.pages.forEach(p => { seeded[p] = g.doc_id }))

      setPreview(json)
      setAssign(seeded)
      setStep('review')
    } catch (err) {
      setError(err.message || 'Could not analyze that file')
      setStep('choose')
    }
  }

  const onPickPdf = (e) => {
    const f = e.target.files?.[0]
    e.target.value = ''   // allow re-picking the same file
    if (!f) return
    if (f.type && f.type !== 'application/pdf') {
      setError('Please choose a PDF file.')
      return
    }
    analyze(f)
  }

  // Live regrouping from the editable map, so the summary and the commit
  // payload always reflect what's on screen rather than the original proposal.
  const groups = useMemo(() => {
    if (!preview) return []
    const byDoc = new Map()
    for (let p = 1; p <= preview.page_count; p++) {
      const id = assign[p]
      if (!id) continue
      if (!byDoc.has(id)) byDoc.set(id, [])
      byDoc.get(id).push(p)
    }
    return [...byDoc.entries()].map(([doc_id, pages]) => ({
      doc_id,
      pages,
      meta: options.find(c => c.doc_id === doc_id),
    }))
  }, [assign, preview, options])

  const unassignedPages = useMemo(() => {
    if (!preview) return []
    const out = []
    for (let p = 1; p <= preview.page_count; p++) if (!assign[p]) out.push(p)
    return out
  }, [assign, preview])

  const commit = async () => {
    if (groups.length === 0) return
    setStep('committing'); setError('')
    try {
      const fd = new FormData()
      fd.append('file', file)
      fd.append('assignments', JSON.stringify(groups.map(g => ({ doc_id: g.doc_id, pages: g.pages }))))
      const res = await fetch('/api/approvals/month-end/commit', { method: 'POST', body: fd })
      const json = await res.json()
      if (!res.ok || json.error) throw new Error(json.error || `HTTP ${res.status}`)
      onUploaded?.(json.uploaded || [])
    } catch (err) {
      setError(err.message || 'Upload failed')
      setStep('review')
    }
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4"
      style={{ backgroundColor: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(2px)' }}
      onClick={() => { if (!busy) onClose?.() }}
    >
      <div
        className="bg-white rounded-2xl shadow-2xl w-full max-w-3xl p-5 space-y-4 max-h-[92vh] overflow-y-auto"
        onClick={e => e.stopPropagation()}
      >
        <div className="flex items-start justify-between gap-3">
          <div className="min-w-0">
            <h2 className="text-lg font-black text-navy">Upload Signed Month-End Docs</h2>
            <p className="text-[12px] text-slate-500">
              {step === 'review'
                ? 'Check the split before it files anything.'
                : 'One combined stack or a single form — both work.'}
            </p>
          </div>
          <button
            type="button"
            onClick={() => { if (!busy) onClose?.() }}
            className="text-slate-400 hover:text-slate-600 text-xl leading-none flex-shrink-0">
            ×
          </button>
        </div>

        {error && (
          <div className="bg-red-50 border border-red-200 text-red-700 text-xs px-3 py-2 rounded-lg">
            {error}
          </div>
        )}

        {step === 'choose' && (
          <>
            <label className="block border-2 border-dashed border-slate-300 rounded-xl p-8 text-center cursor-pointer hover:border-navy transition-colors">
              <input type="file" accept="application/pdf" className="hidden" onChange={onPickPdf} />
              <div className="text-3xl mb-1">📄</div>
              <p className="text-sm font-semibold text-navy">Choose the signed PDF</p>
              <p className="text-[12px] text-slate-500 mt-1">
                Scan the whole signed stack as one PDF, or upload one form at a time (≤ 20 MB)
              </p>
            </label>
            <p className="text-[11px] text-slate-400">
              {candidates.length} document{candidates.length === 1 ? '' : 's'} currently awaiting a
              signed copy. Pages are matched against those, and you confirm the split before
              anything is filed.
            </p>
          </>
        )}

        {step === 'analyzing' && (
          <div className="flex items-center gap-3 py-10 justify-center">
            <div className="w-5 h-5 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
            <span className="text-sm text-slate-600">Reading the pages…</span>
          </div>
        )}

        {step === 'committing' && (
          <div className="flex items-center gap-3 py-10 justify-center">
            <div className="w-5 h-5 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
            <span className="text-sm text-slate-600">Filing into the review queue…</span>
          </div>
        )}

        {step === 'review' && preview && (
          <>
            {(preview.warnings || []).length > 0 && (
              <div className="bg-amber-50 border border-amber-200 text-amber-800 text-xs px-3 py-2 rounded-lg space-y-1">
                {preview.warnings.map((w, i) => <p key={i}>{w}</p>)}
              </div>
            )}

            {/* Summary of what will be filed — derived live from the edits below. */}
            <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 space-y-1.5">
              <p className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">
                Will file {groups.length} document{groups.length === 1 ? '' : 's'}
              </p>
              {groups.length === 0 && (
                <p className="text-xs text-slate-500 italic">
                  Nothing assigned yet — set a document for at least one page below.
                </p>
              )}
              {groups.map(g => (
                <p key={g.doc_id} className="text-xs text-slate-700">
                  <span className="font-semibold">{g.meta ? docLabel(g.meta) : g.doc_id}</span>
                  {' — '}page{g.pages.length === 1 ? '' : 's'} {g.pages.join(', ')}
                </p>
              ))}
              {unassignedPages.length > 0 && (
                <p className="text-xs text-slate-500">
                  Page{unassignedPages.length === 1 ? '' : 's'} {unassignedPages.join(', ')} will be
                  left out.
                </p>
              )}
            </div>

            {/* Per-page assignment. A page-level control is the only view that
                can't misrepresent the split — the summary above is derived from it. */}
            <Document
              file={file}
              loading={<p className="text-xs text-slate-500 py-4">Rendering pages…</p>}
              error={<p className="text-xs text-red-600 py-4">Couldn&apos;t render page previews — you can still assign pages by number.</p>}
              className="space-y-2"
            >
              {Array.from({ length: preview.page_count }, (_, i) => {
                const pageNum = i + 1
                return (
                  <div
                    key={pageNum}
                    className={`flex items-center gap-3 border rounded-lg p-2
                                ${assign[pageNum] ? 'border-slate-200' : 'border-amber-300 bg-amber-50/40'}`}
                  >
                    <div className="flex-shrink-0 w-[70px] overflow-hidden rounded border border-slate-200 bg-white">
                      <Page
                        pageNumber={pageNum}
                        width={70}
                        renderTextLayer={false}
                        renderAnnotationLayer={false}
                      />
                    </div>
                    <div className="min-w-0 flex-1">
                      <p className="text-[11px] font-bold text-slate-500 mb-1">Page {pageNum}</p>
                      <select
                        value={assign[pageNum] || UNASSIGNED}
                        onChange={e => setAssign(a => ({ ...a, [pageNum]: e.target.value }))}
                        className="w-full text-xs border border-slate-300 rounded-lg px-2 py-1.5 bg-white"
                      >
                        <option value={UNASSIGNED}>— skip this page —</option>
                        {options.map(c => (
                          <option key={c.doc_id} value={c.doc_id}>{docLabel(c)}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                )
              })}
            </Document>

            <div className="flex items-center justify-end gap-2 pt-1">
              <button
                type="button"
                onClick={() => { setStep('choose'); setPreview(null); setFile(null); setError('') }}
                className="text-xs font-bold px-3 py-2 rounded-lg bg-slate-100 text-slate-700 hover:bg-slate-200"
              >
                Choose a different file
              </button>
              <button
                type="button"
                onClick={commit}
                disabled={groups.length === 0}
                className="text-xs font-bold px-4 py-2 rounded-lg bg-green-600 text-white
                           hover:bg-green-700 disabled:opacity-50 disabled:cursor-not-allowed"
              >
                Confirm &amp; queue {groups.length || ''}
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  )
}
