import { useEffect, useMemo, useRef, useState } from 'react'
import { Link } from 'react-router-dom'
import { Document, Page, pdfjs } from 'react-pdf'
import 'react-pdf/dist/Page/AnnotationLayer.css'
import 'react-pdf/dist/Page/TextLayer.css'
import PrincipalSignModal from '../components/PrincipalSignModal'
import RowKebab from '../components/RowKebab'
import GenerateDocModal from '../components/GenerateDocModal'
import SignInHoursEditor from '../components/SignInHoursEditor'
import SignInHeaderCard from '../components/SignInHeaderCard'
import ReuploadModal from '../components/ReuploadModal'
import ConfirmModal from '../components/ConfirmModal'
import { usePendingCounts } from '../lib/PendingCountsContext'

// Doc types whose data-driven PDF can be rebuilt from current system data
// and overwritten in place. CP + SI are not yet supported.
const REGENERABLE_TYPES = new Set(['field_report', 'production_log'])
const canRegenerate = (item) => !!item && REGENERABLE_TYPES.has(item.doc_type)

// Manually-uploaded sign-in PDFs (filename ends in _MANUAL.pdf) come
// from the Sign-In tab's "Upload PDF" path. The principal usually
// wet-signed those by hand, so we default the primary approval action
// to Skip Sign-Off and tuck the electronic-overlay action behind a
// kebab in case admin still wants it.
const isManualSignIn = (item) =>
  item?.doc_type === 'signin' &&
  /_MANUAL\.pdf$/i.test(item?.filename || '')

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
  const [processingModalOpen, setProcessingModalOpen] = useState(false)
  const containerRef = useRef(null)
  const [viewerWidth, setViewerWidth] = useState(700)
  // Shared cache feeding the nav badge + the worker-button label. We
  // overwrite these counts after every refresh so the rest of the app
  // sees fresh numbers without a separate poll.
  const { counts, setCount } = usePendingCounts()

  // Fetch pending approvals once on mount + after each approve action
  const refresh = async () => {
    try {
      const res  = await fetch('/api/approvals')
      const data = await res.json()
      if (data.error) throw new Error(data.error)
      const approvalsList = Array.isArray(data.approvals) ? data.approvals : []
      setApprovals(approvalsList)
      setLoadError('')
      // Push the freshest counts into the shared context. approvals_review
      // = needs-review queue length; approved_docs_pending = files in the
      // Approved Docs folder waiting for the worker to pick them up.
      setCount('approvals_review', approvalsList.length)
      if (data.approved_docs_pending !== undefined) {
        setCount('approved_docs_pending', data.approved_docs_pending)
      }
    } catch (err) {
      setLoadError(err.message || 'Failed to load approvals')
      setApprovals([])
    }
  }
  useEffect(() => { refresh() }, [])

  // Manually trigger the same job the 10-min cron runs. Modal-driven
  // for clear sticky feedback (replaces the old 6s auto-dismiss toast).
  const openProcessModal = () => setProcessingModalOpen(true)
  const handleProcessConfirm = async () => {
    const res  = await fetch('/api/tools/process-approved-docs', { method: 'POST' })
    const data = await res.json().catch(() => ({}))
    if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
    const archived = data.archived ?? 0
    const errored  = data.errored  ?? 0
    if (data.skipped) {
      return { message: 'Another run is already in progress — it will finish shortly.' }
    }
    if (archived === 0 && errored === 0) {
      return { message: 'No pending docs to process.' }
    }
    const parts = []
    if (archived > 0) parts.push(`archived ${archived}`)
    if (errored  > 0) parts.push(`${errored} failed → Archive Errors`)
    const msg = `Processed: ${parts.join(', ')}.`
    if (errored > 0) throw new Error(msg)
    return { message: msg }
  }

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

  // Sign-In docs route through PrincipalSignModal; everything else fires
  // the lightweight direct approve.
  const [signingItem, setSigningItem] = useState(null)
  // True while the sign-in hours editor (right pane) has unsaved edits —
  // approving moves the file, so warn before discarding those edits.
  const [signinEditsDirty, setSigninEditsDirty] = useState(false)
  // Controls the Reupload modal (replace the pending PDF with a signed/
  // rescanned version).
  const [reuploadOpen, setReuploadOpen] = useState(false)
  // Certified Payroll: confirm before approving without an electronic
  // signature (guards against pushing a still-unsigned CP through).
  const [cpSkipConfirmOpen, setCpSkipConfirmOpen] = useState(false)
  // Bumped after a reupload to cache-bust the PDF viewer — the reupload
  // replaces content in place (same file_id), so the URL is otherwise
  // unchanged and react-pdf wouldn't re-fetch.
  const [pdfVersion, setPdfVersion] = useState(0)
  // Regenerate flow: confirm dialog + the file_id currently being rebuilt
  // (drives the "Regenerating…" label + a busy banner) + any error.
  const [regenConfirmOpen, setRegenConfirmOpen] = useState(false)
  const [regenBusyId, setRegenBusyId] = useState(null)
  const [regenError, setRegenError] = useState('')

  // Gate any approve action when the hours editor has unsaved edits.
  const confirmDiscardEdits = () =>
    !signinEditsDirty ||
    window.confirm(
      'You have unsaved hour edits. Approving moves the sheet and those ' +
      'edits will be lost unless you Save them first. Approve anyway?'
    )

  // Remove the just-approved item from local state + advance to next in
  // the filtered view. Shared by direct-approve and signed-approve paths.
  const removeApprovedAndAdvance = (fileId) => {
    const idx = filtered.findIndex(a => a.file_id === fileId)
    const nextItem = filtered[idx + 1] || filtered[idx - 1] || null
    setApprovals(prev => (prev || []).filter(a => a.file_id !== fileId))
    setSelectedId(nextItem?.file_id || null)
  }

  const handleApprove = async () => {
    if (!selected || approving) return
    if (!confirmDiscardEdits()) return
    setActionError('')
    // Manually-uploaded sign-in: skip the electronic sign-off (already
    // wet-signed by the principal in person).
    if (isManualSignIn(selected)) {
      return doSkipSignoff()
    }
    // Generated sign-in: open the principal-signature modal.
    if (selected.doc_type === 'signin') {
      setSigningItem(selected)
      return
    }
    // Certified Payroll: admins wet-sign + reupload the signed copy, so the
    // primary action approves WITHOUT a second electronic overlay. Reupload
    // leaves no filename marker (unlike sign-in's _MANUAL), so we can't tell a
    // signed reupload from a still-blank generated CP — a confirm guards against
    // approving an unsigned sheet. Electronic sign-off stays available via the
    // kebab (openCertPayrollSignoff).
    if (selected.doc_type === 'certified_payroll') {
      setCpSkipConfirmOpen(true)
      return
    }
    // Everything else: direct approve.
    setApproving(true)
    const fileId = selected.file_id
    try {
      const res = await fetch(
        `/api/approvals/${encodeURIComponent(fileId)}/approve`,
        { method: 'POST' })
      const data = await res.json()
      if (data.error) throw new Error(data.error)
      removeApprovedAndAdvance(fileId)
    } catch (err) {
      setActionError(err.message || 'Approval failed')
    } finally {
      setApproving(false)
    }
  }

  // Skip Sign-Off: move the file to Approved Docs without overlaying a
  // principal signature. Default action for `_MANUAL` uploads; available
  // via kebab for any sign-in where admin wants to bypass.
  const doSkipSignoff = async () => {
    if (!selected || approving) return
    setActionError('')
    setApproving(true)
    const fileId = selected.file_id
    try {
      const res = await fetch(
        `/api/approvals/${encodeURIComponent(fileId)}/skip-signoff`,
        { method: 'POST' })
      const data = await res.json()
      if (data.error) throw new Error(data.error)
      removeApprovedAndAdvance(fileId)
    } catch (err) {
      setActionError(err.message || 'Skip sign-off failed')
    } finally {
      setApproving(false)
    }
  }

  // Called by PrincipalSignModal after its signed-approve POST succeeds
  const handleSignedApproved = () => {
    if (!signingItem) return
    removeApprovedAndAdvance(signingItem.file_id)
    setSigningItem(null)
  }

  // Certified Payroll electronic sign-off (kebab fallback): opens the same
  // PrincipalSignModal used for sign-ins, bound to the CP signing endpoint.
  const openCertPayrollSignoff = () => {
    if (!selected) return
    setSigningItem({
      ...selected,
      signUrl: `/api/approvals/${encodeURIComponent(selected.file_id)}/approve-cert-payroll`,
    })
  }

  // Called by ReuploadModal after the PDF content is replaced in place.
  // The file_id, name, created date, and list position are unchanged, so
  // we only need to cache-bust the PDF viewer to show the new bytes. The
  // sign-in hours editor stays mounted (same file_id) — its data is
  // unaffected by a PDF swap, so unsaved edits are preserved.
  const handleReuploaded = () => {
    setReuploadOpen(false)
    setPdfVersion(v => v + 1)
  }

  // Snapshot a pending file's "was it overwritten yet?" signature for the
  // regenerate poll. We key on modifiedTime (not md5): the worker's in-place
  // overwrite (Drive.Files.update) ALWAYS bumps modifiedTime — even when the
  // rebuilt PDF is byte-identical (e.g. the admin regenerated without changing
  // any data). md5 alone would miss that case and the poll would falsely
  // time out. Returns '' on failure.
  const fetchFileSig = async (fileId) => {
    try {
      const m = await fetch(`/api/approvals/${encodeURIComponent(fileId)}/meta`)
        .then(r => r.json())
      return (m && m.modified_time) || ''
    } catch { return '' }
  }

  // Poll until the file's modifiedTime differs from the pre-regenerate
  // snapshot, i.e. the worker has overwritten it in place. Returns true on
  // change, false on timeout.
  const pollForOverwrite = async (fileId, beforeSig) => {
    const deadline = Date.now() + 90_000   // ~90s ceiling
    while (Date.now() < deadline) {
      await new Promise(r => setTimeout(r, 2500))
      const sig = await fetchFileSig(fileId)
      if (sig && sig !== beforeSig) return true
    }
    return false
  }

  // Regenerate the selected pending doc (CFR / Production Log) from current
  // system data and overwrite it in place. Asynchronous: queue the rebuild,
  // then poll until the worker refills the same file, then refresh preview.
  const handleRegenerate = async () => {
    setRegenConfirmOpen(false)
    if (!selected || regenBusyId) return
    const fileId = selected.file_id
    setRegenError('')
    setActionError('')
    const beforeSig = await fetchFileSig(fileId)
    setRegenBusyId(fileId)
    try {
      const res = await fetch(
        `/api/approvals/${encodeURIComponent(fileId)}/regenerate`,
        { method: 'POST' })
      const data = await res.json().catch(() => ({}))
      if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
      // beforeSig should be non-empty for a real pending file; if the meta
      // call failed, fall back to a fixed wait before refreshing.
      const changed = beforeSig
        ? await pollForOverwrite(fileId, beforeSig)
        : (await new Promise(r => setTimeout(r, 10_000)), true)
      if (changed) {
        setPdfVersion(v => v + 1)   // re-fetch the overwritten bytes
        refresh()                    // pick up new size/created_at in the list
      } else {
        setRegenError('Still regenerating — the worker may be busy. The preview ' +
          'will refresh on its own shortly, or reselect this doc in a moment.')
      }
    } catch (err) {
      setRegenError(err.message || 'Regenerate failed')
    } finally {
      setRegenBusyId(null)
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
      <div className="flex items-start justify-between gap-4 flex-wrap">
        <div>
          <h1 className="text-2xl font-black text-navy">Approvals</h1>
          <p className="text-sm text-slate-500 mt-1">
            Review and approve generated documents. Approving moves the
            file to <span className="font-mono">✅ Approved Docs</span>, which
            triggers the email + archive job within ~10 minutes.
          </p>
        </div>
        <div className="flex items-center gap-3 flex-shrink-0">
          <button
            type="button"
            onClick={openProcessModal}
            title="Run the same archive + email job the 10-minute cron runs"
            className="text-xs font-bold px-3 py-1.5 rounded-lg
                       bg-navy text-white hover:opacity-90"
          >
            Process Approved Docs now
          </button>
          {/* Number of files waiting in the Approved Docs folder for the
              worker — distinct from the Need Review count above (the All
              pill already surfaces that). Hidden until backend fills the
              field so old proxies don't show a stale "0 pending". */}
          {counts.approved_docs_pending != null && (
            <span
              className="text-xs font-bold text-slate-600"
              title={`${counts.approved_docs_pending} file(s) waiting in the Approved Docs folder for the next processing run`}
            >
              {counts.approved_docs_pending} pending
            </span>
          )}
        </div>
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
                <div className="min-w-0 flex-1">
                  <p className="text-sm font-bold text-slate-800 truncate">
                    {selected.filename}
                  </p>
                  <p className="text-[11px] text-slate-500 mt-0.5 truncate">
                    {docTypeMeta(selected.doc_type).label}
                    {selected.subtitle && <> · {selected.subtitle}</>}
                    <> · {fmtTime(selected.created_at)}</>
                  </p>
                </div>
                <div className="flex flex-wrap items-center gap-2">
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
                    onClick={() => setReuploadOpen(true)}
                    disabled={approving || !!regenBusyId}
                    title="Replace this PDF with a signed/rescanned version"
                    className="text-xs font-bold px-3 py-1.5 rounded-lg
                               bg-slate-100 text-slate-700 hover:bg-slate-200 transition-colors
                               disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    Reupload
                  </button>
                  {canRegenerate(selected) && (
                    <button
                      type="button"
                      onClick={() => setRegenConfirmOpen(true)}
                      disabled={approving || !!regenBusyId}
                      title="Rebuild this document from current system data and replace it in place"
                      className="text-xs font-bold px-3 py-1.5 rounded-lg
                                 bg-blue-100 text-blue-800 hover:bg-blue-200 transition-colors
                                 disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                      {regenBusyId === selected.file_id ? 'Regenerating…' : 'Regenerate'}
                    </button>
                  )}
                  <button
                    type="button"
                    onClick={handleApprove}
                    disabled={approving || !!regenBusyId}
                    className="text-xs font-bold px-4 py-1.5 rounded-lg
                               bg-green-600 text-white hover:bg-green-700 transition-colors
                               disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    {approving
                      ? 'Approving…'
                      : (isManualSignIn(selected) || selected.doc_type === 'certified_payroll')
                          ? 'Approve (skip sign-off)'
                          : selected.doc_type === 'signin'
                              ? 'Approve & Sign'
                              : 'Approve'}
                  </button>
                  {/* Sign-in overrides — kebab keeps the rare alternate
                      action available without cluttering the row. */}
                  {selected.doc_type === 'signin' && !approving && (
                    <RowKebab items={
                      isManualSignIn(selected)
                        ? [{
                            label:   'Approve with sign-off',
                            onClick: () => { if (confirmDiscardEdits()) setSigningItem(selected) },
                          }]
                        : [{
                            label:   'Skip sign-off',
                            onClick: () => { if (confirmDiscardEdits()) doSkipSignoff() },
                          }]
                    } />
                  )}
                  {/* Certified Payroll override — primary action skips sign-off
                      (admins wet-sign + reupload); the electronic overlay stays
                      available here for the occasional in-app signature. */}
                  {selected.doc_type === 'certified_payroll' && !approving && (
                    <RowKebab items={[{
                      label:   'Approve with sign-off',
                      onClick: openCertPayrollSignoff,
                    }]} />
                  )}
                </div>
              </div>

              {actionError && (
                <div className="bg-red-50 border border-red-200 text-red-700 text-xs
                                px-3 py-2 rounded-lg mb-3">
                  {actionError}
                </div>
              )}

              {regenBusyId === selected.file_id && (
                <div className="bg-blue-50 border border-blue-200 text-blue-800 text-xs
                                px-3 py-2 rounded-lg mb-3 flex items-center gap-2">
                  <span className="inline-block w-3 h-3 border-2 border-blue-400
                                   border-t-transparent rounded-full animate-spin" />
                  Regenerating from current data — the preview updates automatically
                  once the worker refills it (usually a few seconds).
                </div>
              )}

              {regenError && (
                <div className="bg-amber-50 border border-amber-200 text-amber-800 text-xs
                                px-3 py-2 rounded-lg mb-3">
                  {regenError}
                </div>
              )}

              {/* Read-only sign-in sheet header — mirrors the Sign-In tab
                  header so an admin can hand-fill a printed sheet's header
                  when the crew left it blank. Sign-in sheets only. */}
              {selected.doc_type === 'signin' && (
                <SignInHeaderCard
                  key={`hdr-${selected.file_id}`}
                  fileId={selected.file_id}
                  filename={selected.filename}
                />
              )}

              {/* Editable recorded-hours table — sign-in sheets only.
                  Other doc types keep the plain read-only PDF view. */}
              {selected.doc_type === 'signin' && (
                <SignInHoursEditor
                  key={selected.file_id}
                  fileId={selected.file_id}
                  filename={selected.filename}
                  onDirtyChange={setSigninEditsDirty}
                />
              )}

              <PDFViewer fileId={selected.file_id} width={viewerWidth} version={pdfVersion} />
            </>
          )}
        </section>
      </div>

      <GenerateDocModal
        open={processingModalOpen}
        title="Process Approved Documents"
        description={
          <span>
            Run the same archive + email job the 10-minute cron runs. Safe to run alongside the cron — script-wide lock prevents double-processing.
          </span>
        }
        confirmLabel="Run Now"
        onConfirm={async () => {
          const result = await handleProcessConfirm()
          await refresh()
          return result
        }}
        onClose={() => setProcessingModalOpen(false)}
      />

      {/* Approval-with-signature modal — used for both Sign-In sheets and
          Certified Payroll. CP pre-binds its own signUrl on the signingItem;
          sign-ins fall back to the default /approve-signin endpoint. */}
      {signingItem && (
        <PrincipalSignModal
          filename={signingItem.filename}
          signUrl={signingItem.signUrl ||
            `/api/approvals/${encodeURIComponent(signingItem.file_id)}/approve-signin`}
          onCancel={() => setSigningItem(null)}
          onSigned={handleSignedApproved}
        />
      )}

      {/* Reupload modal — replace the pending PDF with a signed/rescanned
          version. Sign-ins get scan + PDF; other doc types get PDF only. */}
      {reuploadOpen && selected && (
        <ReuploadModal
          fileId={selected.file_id}
          filename={selected.filename}
          allowScan={selected.doc_type === 'signin'}
          onClose={() => setReuploadOpen(false)}
          onReuploaded={handleReuploaded}
        />
      )}

      {/* Certified Payroll approve-without-sign-off confirmation. Reupload
          keeps the same filename, so there's no marker proving the PDF was
          signed — this guards against approving a still-blank CP. */}
      {cpSkipConfirmOpen && selected && (
        <ConfirmModal
          title="Approve without electronic sign-off?"
          message={
            `This moves the Certified Payroll to ✅ Approved Docs with no ` +
            `electronic signature overlay. Make sure the signed copy has been ` +
            `reuploaded first. To add an in-app signature instead, cancel and use ` +
            `“Approve with sign-off” from the ⋮ menu.`
          }
          confirmLabel="Approve (skip sign-off)"
          cancelLabel="Cancel"
          onConfirm={() => { setCpSkipConfirmOpen(false); doSkipSignoff() }}
          onCancel={() => setCpSkipConfirmOpen(false)}
        />
      )}

      {/* Regenerate confirmation — rebuilds the doc from current system
          data and overwrites it in place. */}
      {regenConfirmOpen && selected && (
        <ConfirmModal
          title="Regenerate this document?"
          message={
            `This rebuilds the ${docTypeMeta(selected.doc_type).label} from the current ` +
            `system data and replaces the reviewed PDF in place (same file, same spot in ` +
            `the queue). Make sure you've saved your data corrections first.`
          }
          confirmLabel="Regenerate"
          cancelLabel="Cancel"
          onConfirm={handleRegenerate}
          onCancel={() => setRegenConfirmOpen(false)}
        />
      )}
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
function PDFViewer({ fileId, width, version = 0 }) {
  const [numPages, setNumPages] = useState(null)
  const [loadErr, setLoadErr]   = useState(null)

  // react-pdf's `file` prop. Using the URL directly lets pdf.js stream
  // the document — faster than fetching the whole blob first. `version`
  // is a cache-buster bumped after a reupload (content replaced in place
  // under the same file_id) so react-pdf re-fetches the new bytes.
  const file = useMemo(() => ({
    url: `/api/approvals/${encodeURIComponent(fileId)}/pdf${version ? `?v=${version}` : ''}`,
  }), [fileId, version])

  // Reset error / page count whenever the document changes
  useEffect(() => {
    setLoadErr(null)
    setNumPages(null)
  }, [fileId, version])

  return (
    <div className="flex-1 min-h-0 overflow-auto bg-slate-50 rounded-lg border border-slate-200">
      {loadErr && (
        <div className="p-6 text-sm text-red-700">
          Couldn&apos;t render this PDF: {loadErr.message || String(loadErr)}
        </div>
      )}
      {!loadErr && (
        <Document
          key={`${fileId}:${version}`}
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
