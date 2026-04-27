import { useEffect, useMemo, useRef, useState } from 'react'
import { Link } from 'react-router-dom'
import SignaturePad from '../components/SignaturePad'

// ── Constants ──────────────────────────────────────────────────
// Mirrors the dropdown validation on Daily Sign-In Data → Classification.
// Keep this in sync with FieldReport.jsx and the sheet validation, or
// submissions will fail with "Invalid Entry".
const CLASSIFICATIONS = ['LP', 'SAT']

// ── Helpers ────────────────────────────────────────────────────
const newCrew = () => ({
  name: '',
  classification: CLASSIFICATIONS[0],
  timeIn: '',
  timeOut: '',
  hours: '',
  overtime: '',
  signatureIn: null,
  signatureOut: null,
})

// Day-of-week-aware OT calc. Sat/Sun → all OT; weekday over 8 → OT.
// workDateIso is the SHIFT START date as YYYY-MM-DD. Constructed from
// local parts so it doesn't UTC-shift the day-of-week.
const calcHours = (tin, tout, workDateIso) => {
  if (!tin || !tout) return { hours: '', overtime: '' }
  const [ih, im] = tin.split(':').map(Number)
  const [oh, om] = tout.split(':').map(Number)
  const mins = (oh * 60 + om) - (ih * 60 + im)
  if (mins <= 0) return { hours: '', overtime: '' }   // cross-midnight handled in Phase 3
  const hrs = mins / 60

  let isWeekend = false
  const dm = String(workDateIso || '').match(/^(\d{4})-(\d{2})-(\d{2})/)
  if (dm) {
    const dow = new Date(Number(dm[1]), Number(dm[2]) - 1, Number(dm[3])).getDay()
    isWeekend = (dow === 0 || dow === 6)
  }
  const ot = isWeekend ? hrs : Math.max(0, hrs - 8)
  return { hours: hrs.toFixed(2), overtime: ot.toFixed(2) }
}

// Pretty date for the queue cards: "Apr 26 (Sat)"
const prettyQueueDate = (iso) => {
  const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/)
  if (!m) return String(iso)
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]))
  return d.toLocaleDateString(undefined, {
    month: 'short', day: 'numeric', weekday: 'short',
  })
}

// ── Field wrapper (same pattern as FieldReport) ────────────────
function Field({ label, required, hint, children }) {
  return (
    <div className="space-y-1">
      <label className="field-label">
        {label}{required && <span className="text-red-400 ml-0.5">*</span>}
      </label>
      {children}
      {hint && <p className="text-[11px] text-slate-400">{hint}</p>}
    </div>
  )
}

// ── Crew row (with employee dropdown + sig pads) ───────────────
function CrewRow({ idx, data, employees, onChange, onRemove, workDate }) {
  const handleTime = (field, val) => {
    const tin  = field === 'timeIn'  ? val : data.timeIn
    const tout = field === 'timeOut' ? val : data.timeOut
    onChange(idx, { ...data, [field]: val, ...calcHours(tin, tout, workDate) })
  }

  return (
    <div className="border border-slate-200 rounded-xl p-4 space-y-4 bg-slate-50/40">
      <div className="flex justify-between items-center">
        <span className="text-xs font-bold text-navy uppercase tracking-wider">
          Crew Member #{idx + 1}
        </span>
        <button
          type="button"
          onClick={() => onRemove(idx)}
          className="text-red-400 hover:text-red-600 text-xl leading-none">
          ×
        </button>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <Field label="Employee" required>
          <select
            value={data.name}
            onChange={e => onChange(idx, { ...data, name: e.target.value })}
            className="field-input">
            <option value="">— Select —</option>
            {employees.map(emp => (
              <option key={emp.name} value={emp.name}>{emp.name}</option>
            ))}
          </select>
        </Field>
        <Field label="Classification" required>
          <select
            value={data.classification}
            onChange={e => onChange(idx, { ...data, classification: e.target.value })}
            className="field-input">
            {CLASSIFICATIONS.map(c => <option key={c}>{c}</option>)}
          </select>
        </Field>
      </div>

      <div className="grid grid-cols-2 gap-3">
        <Field label="Time In" required>
          <input
            type="time"
            value={data.timeIn}
            onChange={e => handleTime('timeIn', e.target.value)}
            className="field-input" />
        </Field>
        <Field label="Time Out" required>
          <input
            type="time"
            value={data.timeOut}
            onChange={e => handleTime('timeOut', e.target.value)}
            className="field-input" />
        </Field>
      </div>

      {data.hours && (
        <div className="flex gap-4 bg-white rounded-lg px-3 py-2 border border-slate-200">
          <div>
            <span className="text-[10px] text-slate-400 uppercase tracking-wider font-bold">
              Reg. Hours
            </span>
            <p className="text-sm font-bold text-navy">{data.hours}</p>
          </div>
          <div className="border-l border-slate-200 pl-4">
            <span className="text-[10px] text-slate-400 uppercase tracking-wider font-bold">
              OT Hours
            </span>
            <p className={`text-sm font-bold ${parseFloat(data.overtime) > 0 ? 'text-orange-600' : 'text-slate-400'}`}>
              {data.overtime || '0.00'}
            </p>
          </div>
        </div>
      )}

      <div className="pt-1 border-t border-slate-200 space-y-3">
        <p className="text-[11px] text-slate-400 font-medium">
          Employee signs below — added to the Sign-In Log
        </p>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <SignaturePad
            label="Time-In Signature"
            onChange={dataUrl => onChange(idx, { ...data, signatureIn: dataUrl })}
          />
          <SignaturePad
            label="Time-Out Signature"
            onChange={dataUrl => onChange(idx, { ...data, signatureOut: dataUrl })}
          />
        </div>
      </div>
    </div>
  )
}

// ── Queue card (left pane) ─────────────────────────────────────
function QueueCard({ entry, selected, onSelect }) {
  return (
    <button
      onClick={() => onSelect(entry.queue_id)}
      className={`w-full text-left p-3 border-b border-slate-200 transition-colors
        ${selected ? 'bg-navy/5 border-l-4 border-l-navy' : 'bg-white hover:bg-slate-50'}`}>
      <div className="text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-1">
        {prettyQueueDate(entry.date)}
      </div>
      <div className="text-sm font-semibold text-navy">
        {entry.contract_number} · {entry.borough}
      </div>
      {entry.contractor && (
        <div className="text-[12px] text-slate-500 mt-0.5">{entry.contractor}</div>
      )}
      <div className="text-[12px] text-slate-700 mt-1.5">
        <span className="font-semibold">{entry.wos.length}</span>{' '}
        {entry.wos.length === 1 ? 'WO' : 'WOs'}
        <span className="text-slate-400"> · </span>
        {entry.wos.slice(0, 2).map(w => w.id).join(', ')}
        {entry.wos.length > 2 && (
          <span className="text-slate-400"> +{entry.wos.length - 2}</span>
        )}
      </div>
    </button>
  )
}

// ── Empty state ────────────────────────────────────────────────
function CaughtUp() {
  return (
    <div className="flex flex-col items-center justify-center py-16 px-6 text-center">
      <div className="w-14 h-14 rounded-full bg-emerald-100 text-emerald-600 flex items-center justify-center text-2xl mb-3">
        ✓
      </div>
      <h2 className="text-lg font-bold text-navy mb-1">All caught up</h2>
      <p className="text-sm text-slate-500 max-w-xs">
        No pending sign-ins. New entries appear here as crews submit Field Reports.
      </p>
      <Link to="/" className="btn-ghost mt-4">Back to Dashboard</Link>
    </div>
  )
}

// ── Toast ──────────────────────────────────────────────────────
function Toast({ message, kind = 'success' }) {
  if (!message) return null
  const palette = kind === 'error'
    ? 'bg-red-600 text-white'
    : 'bg-emerald-600 text-white'
  return (
    <div className={`fixed bottom-6 left-1/2 -translate-x-1/2 px-4 py-2.5 rounded-lg shadow-lg text-sm font-semibold z-50 ${palette}`}>
      {message}
    </div>
  )
}

// ── Main page ──────────────────────────────────────────────────
export default function SignIn() {
  const [queue, setQueue] = useState(null)            // null = loading
  const [loadError, setLoadError] = useState('')
  const [selectedId, setSelectedId] = useState(null)
  const [employees, setEmployees] = useState([])
  // Per-queue draft state — switching between entries doesn't wipe progress.
  const [drafts, setDrafts] = useState(() => new Map())
  const [submitting, setSubmitting] = useState(false)
  const [toast, setToast] = useState({ message: '', kind: 'success' })
  const toastTimer = useRef(null)

  // ── Load queue + employees on mount ──────────────────────────
  useEffect(() => {
    let cancelled = false
    Promise.all([
      fetch('/api/signin-queue').then(r => r.json()),
      fetch('/api/employees').then(r => r.json()),
    ])
      .then(([qData, eData]) => {
        if (cancelled) return
        if (qData.error) {
          setLoadError(qData.error)
          setQueue([])
          return
        }
        const list = qData.queue || []
        setQueue(list)
        setEmployees(eData.employees || [])
        // Default selection: first entry in the queue
        if (list.length > 0) setSelectedId(list[0].queue_id)
      })
      .catch(err => {
        if (cancelled) return
        setLoadError(err.message || 'Failed to load sign-in queue')
        setQueue([])
      })
    return () => { cancelled = true }
  }, [])

  // ── Toast helper ─────────────────────────────────────────────
  const showToast = (message, kind = 'success') => {
    setToast({ message, kind })
    if (toastTimer.current) clearTimeout(toastTimer.current)
    toastTimer.current = setTimeout(() => setToast({ message: '', kind: 'success' }), 3500)
  }

  // ── Selected queue entry ─────────────────────────────────────
  const selected = useMemo(
    () => (queue || []).find(q => q.queue_id === selectedId) || null,
    [queue, selectedId]
  )

  // Get-or-create the draft for a queue entry. Draft fields cover the
  // fully-entered form so switching back to an in-progress entry returns
  // exactly to where the user was.
  const getDraft = (entry) => {
    if (!entry) return null
    if (drafts.has(entry.queue_id)) return drafts.get(entry.queue_id)
    const fresh = {
      crew: [newCrew()],
      contractorName: '',
      contractorTitle: '',
      contractorSignature: null,
    }
    setDrafts(prev => new Map(prev).set(entry.queue_id, fresh))
    return fresh
  }

  const updateDraft = (queueId, partial) => {
    setDrafts(prev => {
      const next = new Map(prev)
      const cur = next.get(queueId) || { crew: [newCrew()], contractorName: '', contractorTitle: '', contractorSignature: null }
      next.set(queueId, { ...cur, ...partial })
      return next
    })
  }

  const draft = selected ? getDraft(selected) : null

  // ── Crew handlers (operate on the selected entry's draft) ────
  const updateCrewMember = (idx, member) => {
    if (!selected || !draft) return
    const next = draft.crew.map((m, i) => (i === idx ? member : m))
    updateDraft(selected.queue_id, { crew: next })
  }
  const addCrewMember = () => {
    if (!selected || !draft) return
    updateDraft(selected.queue_id, { crew: [...draft.crew, newCrew()] })
  }
  const removeCrewMember = (idx) => {
    if (!selected || !draft) return
    const next = draft.crew.filter((_, i) => i !== idx)
    updateDraft(selected.queue_id, { crew: next.length ? next : [newCrew()] })
  }

  // ── Submit ───────────────────────────────────────────────────
  const handleSubmit = async () => {
    if (!selected || !draft || submitting) return

    // Light validation — server enforces the rest.
    const crewClean = draft.crew.filter(m => m.name && m.name.trim())
    if (crewClean.length === 0) {
      showToast('Add at least one crew member with a name', 'error')
      return
    }
    for (const m of crewClean) {
      if (!m.timeIn || !m.timeOut) {
        showToast(`Crew member "${m.name}" needs Time In and Time Out`, 'error')
        return
      }
    }

    setSubmitting(true)
    try {
      const res = await fetch('/api/signin', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          queue_id:        selected.queue_id,
          date:            selected.date,
          contract_number: selected.contract_number,
          borough:         selected.borough,
          contractor:      selected.contractor,
          project_name:    selected.project_name,
          contract_id:     selected.contract_id,
          wo_ids:          selected.wos.map(w => w.id),
          crew: crewClean.map(m => ({
            name:           m.name,
            classification: m.classification,
            time_in:        m.timeIn,
            time_out:       m.timeOut,
            hours:          m.hours,
            overtime:       m.overtime,
            sig_in_b64:     m.signatureIn  || '',
            sig_out_b64:    m.signatureOut || '',
          })),
          contractor_name:            draft.contractorName,
          contractor_title:           draft.contractorTitle,
          contractor_signature_b64:   draft.contractorSignature || '',
          date_signed:                '',
          source: 'generated',
        }),
      })
      const json = await res.json()
      if (!res.ok || json.error) throw new Error(json.error || `HTTP ${res.status}`)

      // Drop submitted entry, auto-advance
      const remaining = (queue || []).filter(q => q.queue_id !== selected.queue_id)
      setDrafts(prev => {
        const next = new Map(prev)
        next.delete(selected.queue_id)
        return next
      })
      setQueue(remaining)
      const idx = (queue || []).findIndex(q => q.queue_id === selected.queue_id)
      const nextEntry = remaining[idx] || remaining[idx - 1] || null
      setSelectedId(nextEntry ? nextEntry.queue_id : null)

      showToast(`Sign-in submitted for ${selected.contract_number}`, 'success')
    } catch (err) {
      showToast(err.message || 'Submit failed', 'error')
    } finally {
      setSubmitting(false)
    }
  }

  // ── Render ───────────────────────────────────────────────────
  if (queue === null) {
    return (
      <div className="max-w-6xl mx-auto px-4 py-8">
        <div className="card p-8 text-center text-slate-500">Loading sign-in queue…</div>
      </div>
    )
  }

  if (loadError) {
    return (
      <div className="max-w-6xl mx-auto px-4 py-8">
        <div className="card p-6 border-red-200 bg-red-50">
          <p className="text-sm text-red-700 font-semibold mb-1">
            Failed to load sign-in queue
          </p>
          <p className="text-xs text-red-600">{loadError}</p>
        </div>
      </div>
    )
  }

  if (queue.length === 0) {
    return (
      <div className="max-w-6xl mx-auto px-4 py-8">
        <div className="card"><CaughtUp /></div>
      </div>
    )
  }

  return (
    <div className="max-w-6xl mx-auto px-4 py-6">
      <div className="flex items-baseline justify-between mb-4">
        <h1 className="text-xl font-bold text-navy">Sign-In</h1>
        <span className="text-xs text-slate-500">
          {queue.length} pending
        </span>
      </div>

      <div className="grid md:grid-cols-[320px_1fr] gap-4">
        {/* Queue list */}
        <div className="card overflow-hidden md:max-h-[78vh] md:overflow-y-auto">
          {queue.map(entry => (
            <QueueCard
              key={entry.queue_id}
              entry={entry}
              selected={entry.queue_id === selectedId}
              onSelect={setSelectedId}
            />
          ))}
        </div>

        {/* Entry form */}
        {selected && draft ? (
          <div className="space-y-4">
            {/* Header card — read-only context */}
            <div className="card p-4 space-y-2">
              <div className="grid grid-cols-2 gap-3 text-sm">
                <div>
                  <span className="field-label">Date of Work</span>
                  <p className="font-semibold text-navy">{prettyQueueDate(selected.date)}</p>
                  <p className="text-[11px] text-slate-400">
                    Shift start (auto-detected from Field Report)
                  </p>
                </div>
                <div>
                  <span className="field-label">Contract</span>
                  <p className="font-semibold text-navy">
                    {selected.contract_number} · {selected.borough}
                  </p>
                  {selected.contract_id && (
                    <p className="text-[11px] text-slate-400">{selected.contract_id}</p>
                  )}
                </div>
              </div>
              <div>
                <span className="field-label">Contractor</span>
                <p className="text-sm text-slate-700">{selected.contractor || '—'}</p>
              </div>
              {selected.project_name && (
                <div>
                  <span className="field-label">Project</span>
                  <p className="text-sm text-slate-700">{selected.project_name}</p>
                </div>
              )}
              <div>
                <span className="field-label">Work Orders ({selected.wos.length})</span>
                <ul className="text-sm text-slate-700 space-y-0.5 mt-0.5">
                  {selected.wos.map(w => (
                    <li key={w.id} className="flex gap-2">
                      <span className="font-mono font-semibold">{w.id}</span>
                      {w.location && (
                        <span className="text-slate-500">— {w.location}</span>
                      )}
                    </li>
                  ))}
                </ul>
              </div>
            </div>

            {/* Crew */}
            <div className="card p-4 space-y-4">
              <div className="flex items-center justify-between">
                <p className="section-label !mb-0">
                  Crew & Signatures ({draft.crew.length})
                </p>
              </div>
              <div className="space-y-3">
                {draft.crew.map((m, i) => (
                  <CrewRow
                    key={i}
                    idx={i}
                    data={m}
                    employees={employees}
                    onChange={updateCrewMember}
                    onRemove={removeCrewMember}
                    workDate={selected.date}
                  />
                ))}
              </div>
              <button
                type="button"
                onClick={addCrewMember}
                className="btn-outline w-full">
                + Add crew member
              </button>
            </div>

            {/* Contractor sign-off (crew leader at bottom of form) */}
            <div className="card p-4 space-y-3">
              <p className="section-label">Crew Leader Sign-Off</p>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <Field label="Printed Name" required>
                  <input
                    type="text"
                    value={draft.contractorName}
                    onChange={e => updateDraft(selected.queue_id, { contractorName: e.target.value })}
                    className="field-input" />
                </Field>
                <Field label="Title">
                  <input
                    type="text"
                    value={draft.contractorTitle}
                    onChange={e => updateDraft(selected.queue_id, { contractorTitle: e.target.value })}
                    className="field-input" />
                </Field>
              </div>
              <SignaturePad
                label="Signature"
                onChange={dataUrl => updateDraft(selected.queue_id, { contractorSignature: dataUrl })}
              />
            </div>

            <div className="sticky bottom-4 z-10">
              <button
                onClick={handleSubmit}
                disabled={submitting}
                className="btn-primary w-full text-base">
                {submitting ? 'Submitting…' : 'Submit Sign-In'}
              </button>
            </div>
          </div>
        ) : (
          <div className="card p-8 text-center text-slate-500">
            Select an entry from the queue
          </div>
        )}
      </div>

      <Toast message={toast.message} kind={toast.kind} />
    </div>
  )
}
