import { useEffect, useMemo, useRef, useState } from 'react'
import { Link } from 'react-router-dom'
import { Document, Page, pdfjs } from 'react-pdf'
import 'react-pdf/dist/Page/AnnotationLayer.css'
import 'react-pdf/dist/Page/TextLayer.css'
import SignaturePad from '../components/SignaturePad'
import RowKebab     from '../components/RowKebab'
import ConfirmModal from '../components/ConfirmModal'
import { opDayFromIsoTime } from '../lib/dateOps'

// Mirror the Approvals page wiring — same pdf.js worker URL, version
// pinned to the bundled react-pdf so we don't drift if Vite swaps it.
pdfjs.GlobalWorkerOptions.workerSrc =
  `https://unpkg.com/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`

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
  // Set when the upload-flow OCR returned a name but it didn't match the
  // Employee Registry exactly. Surfaces below the dropdown so the user
  // can see what Claude read and pick the matching employee.
  parsedNameRaw: '',
})

// Normalize a name for fuzzy matching — lowercase, collapse spaces,
// drop punctuation. Used to map OCR'd names against the Employee
// Registry without tripping on "John A. Smith" vs "John Smith" etc.
const normName = (s) => String(s || '')
  .toLowerCase()
  .replace(/[.,]/g, '')
  .replace(/\s+/g, ' ')
  .trim()

// Day-of-week-aware OT calc. Sat/Sun → all OT; weekday over 8 → OT.
// workDateIso is the SHIFT START date as YYYY-MM-DD. Constructed from
// local parts so it doesn't UTC-shift the day-of-week.
//
// Cross-midnight: if Time Out <= Time In we assume the shift rolled
// over to the next calendar day and add 24h. Hours are still bucketed
// under the start day for OT purposes — a Fri-night → Sat-morning
// shift counts as 8h Friday-rate.
const calcHours = (tin, tout, workDateIso) => {
  if (!tin || !tout) return { hours: '', overtime: '' }
  const [ih, im] = tin.split(':').map(Number)
  const [oh, om] = tout.split(':').map(Number)
  let mins = (oh * 60 + om) - (ih * 60 + im)
  if (mins <= 0) mins += 24 * 60
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
// `showSignatures` is false for the Upload tab — uploaded scans
// already carry the crew's signatures on paper, so we only need to
// confirm the parsed identity / classification / time-in-out values.
function CrewRow({ idx, data, employees, onChange, onRemove, workDate, showSignatures = true }) {
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
        <Field
          label="Employee"
          required
          hint={data.parsedNameRaw && !data.name
            ? `Parsed as "${data.parsedNameRaw}" — pick the matching employee, or add them to the Registry first.`
            : undefined}>
          <select
            value={data.name}
            onChange={e => onChange(idx, {
              ...data, name: e.target.value,
              // Once user makes a manual pick, drop the raw-parse hint.
              parsedNameRaw: '',
            })}
            className={`field-input ${data.parsedNameRaw && !data.name ? 'border-amber-300 bg-amber-50/50' : ''}`}>
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
              Hours (this sheet)
            </span>
            <p className="text-sm font-bold text-navy">{data.hours}</p>
          </div>
        </div>
      )}

      {showSignatures && (
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
      )}
    </div>
  )
}

// ── Inline PDF preview (used in Upload mode after parse) ───────
// Renders the uploaded handwritten sign-in sheet at fit-width inside
// a scrollable container so the user can verify the parsed fields
// without leaving the form.
function UploadPreview({ url }) {
  const [numPages, setNumPages] = useState(null)
  const [width, setWidth]       = useState(640)
  const wrapRef = useRef(null)

  useEffect(() => {
    const el = wrapRef.current
    if (!el) return
    const ro = new ResizeObserver(entries => {
      const w = entries[0]?.contentRect?.width || 640
      setWidth(Math.max(280, Math.floor(w) - 16))
    })
    ro.observe(el)
    return () => ro.disconnect()
  }, [])

  return (
    <div ref={wrapRef} className="w-full">
      <Document
        file={url}
        onLoadSuccess={({ numPages }) => setNumPages(numPages)}
        loading={<p className="text-xs text-slate-500 p-3">Rendering preview…</p>}
        className="flex flex-col items-center gap-2"
      >
        {Array.from({ length: numPages || 0 }, (_, i) => (
          <Page
            key={i + 1}
            pageNumber={i + 1}
            width={width}
            renderTextLayer={false}
            renderAnnotationLayer={false}
            className="shadow-sm border border-slate-200 bg-white"
          />
        ))}
      </Document>
    </div>
  )
}

// ── Shift Totals panel ─────────────────────────────────────────
// Per-employee summary across THIS sign-in plus any other Daily
// Sign-In Data rows already on this operational day. Applies the OT
// rule (Sat/Sun all OT, weekday over 8 OT) to the full total so the
// user can see whether multi-contract overnight work has pushed
// anyone into overtime.
function ShiftTotalsPanel({ totals, dateLabel }) {
  if (!totals.length) return null
  const anyOther = totals.some(t => t.otherSheets > 0)
  return (
    <div className="card p-4 space-y-3">
      <p className="section-label !mb-0">
        Shift Totals · {dateLabel}
      </p>
      <p className="text-[11px] text-slate-400">
        ST/OT shown is per-employee across ALL sign-ins on this date — the
        actual values cert payroll will see once this sheet posts.
      </p>
      <div className="space-y-2">
        <div className="grid grid-cols-12 gap-2 text-[10px] uppercase tracking-wider font-bold text-slate-400 px-1">
          <span className="col-span-4">Employee</span>
          <span className="col-span-2 text-right">This sheet</span>
          <span className="col-span-2 text-right">Other</span>
          <span className="col-span-2 text-right">Total</span>
          <span className="col-span-1 text-right">ST</span>
          <span className="col-span-1 text-right">OT</span>
        </div>
        {totals.map(t => (
          <div key={t.name}
            className="grid grid-cols-12 gap-2 text-sm py-1.5 px-1 border-t border-slate-100 items-baseline">
            <span className="col-span-4 font-semibold text-navy truncate">{t.name}</span>
            <span className="col-span-2 text-right text-slate-700">{t.thisSheet.toFixed(2)}</span>
            <span className={`col-span-2 text-right ${t.otherSheets > 0 ? 'text-slate-700' : 'text-slate-300'}`}>
              {t.otherSheets.toFixed(2)}
            </span>
            <span className="col-span-2 text-right font-bold text-navy">{t.total.toFixed(2)}</span>
            <span className="col-span-1 text-right text-slate-700">{t.st.toFixed(2)}</span>
            <span className={`col-span-1 text-right font-bold ${t.ot > 0 ? 'text-orange-600' : 'text-slate-300'}`}>
              {t.ot.toFixed(2)}
            </span>
          </div>
        ))}
      </div>
      {!anyOther && (
        <p className="text-[11px] text-slate-400">
          No other sign-ins on this date yet — totals match this sheet.
        </p>
      )}
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
      // Idempotency token — sent with every submit attempt for this
      // draft. The server caches the response keyed by this id, so a
      // retry after a perceived error (network blip, slow response,
      // etc.) returns the cached success instead of duplicating Daily
      // Sign-In Data rows. New ID generated per draft, not per attempt.
      submitId: (typeof crypto !== 'undefined' && crypto.randomUUID)
        ? crypto.randomUUID()
        : (Math.random().toString(36).slice(2) + Date.now().toString(36)),
      // Default = the queue entry's date (= the WO's Field Report Date of
      // Work). Editable behind a warning modal — see the kebab in the
      // header card. Drives OT calculation: a Fri-night → Sat-morning
      // shift kept as start=Friday gives Friday OT rules.
      shiftStartDate: entry.date,
      shiftDateEdited: false,
      // Upload-flow state. Populated after Claude Vision parses an
      // uploaded scan; the original file's bytes ride along to submit
      // so Apps Script archives the wet-signed paper instead of a
      // generated PDF.
      mode: 'generate',                // 'generate' | 'upload'
      uploadFilename: '',
      uploadMimeType: '',
      uploadDataB64:  '',
      parseStatus:    'idle',          // 'idle' | 'parsing' | 'parsed' | 'error'
      parseError:     '',
    }
    setDrafts(prev => new Map(prev).set(entry.queue_id, fresh))
    return fresh
  }

  const updateDraft = (queueId, partial) => {
    setDrafts(prev => {
      const next = new Map(prev)
      const cur = next.get(queueId) || {
        crew: [newCrew()],
        shiftStartDate: '', shiftDateEdited: false,
        mode: 'generate', uploadFilename: '', uploadMimeType: '',
        uploadDataB64: '', parseStatus: 'idle', parseError: '',
      }
      next.set(queueId, { ...cur, ...partial })
      return next
    })
  }

  // ── Upload + parse handler ───────────────────────────────────
  // Sends the file to Claude Vision via the Express proxy and merges the
  // returned crew rows into the draft so the user can review/correct
  // them in the same form they'd use to generate from scratch.
  const handleUploadFile = async (file) => {
    if (!selected || !file) return
    if (!file.type || file.type !== 'application/pdf') {
      showToast('Upload must be a PDF', 'error')
      return
    }
    updateDraft(selected.queue_id, { parseStatus: 'parsing', parseError: '' })

    try {
      const fd = new FormData()
      fd.append('file', file)
      const res = await fetch('/api/signin-queue/parse-upload', { method: 'POST', body: fd })
      const json = await res.json()
      if (!res.ok || json.error) throw new Error(json.error || `HTTP ${res.status}`)

      const parsed = json.parsed || {}
      const crewIn = Array.isArray(parsed.crew) ? parsed.crew : []
      // Build a normalized index of registry employees so OCR'd names
      // like "Yanni Angelides" still match a registry entry like
      // "YANNI ANGELIDES" or "Yanni  Angelides ". If no match, we keep
      // the raw OCR'd name on the row so the user can see what Claude
      // read while picking the correct employee from the dropdown.
      const normIdx = {}
      employees.forEach(e => { normIdx[normName(e.name)] = e.name })

      const crewMembers = (crewIn.length ? crewIn : [{}]).map(c => {
        const cls = (c.classification === 'LP' || c.classification === 'SAT')
          ? c.classification : CLASSIFICATIONS[0]
        const m = newCrew()
        const rawName = String(c.name || '').trim()
        const matched = rawName ? normIdx[normName(rawName)] : ''
        m.name           = matched || ''
        m.parsedNameRaw  = matched ? '' : rawName
        m.classification = cls
        m.timeIn         = c.time_in  || ''
        m.timeOut        = c.time_out || ''
        const calc = calcHours(m.timeIn, m.timeOut, effectiveDate)
        m.hours    = calc.hours
        m.overtime = calc.overtime
        return m
      })

      updateDraft(selected.queue_id, {
        parseStatus:     'parsed',
        parseError:      '',
        crew:            crewMembers,
        uploadFilename:  json.upload?.filename  || file.name,
        uploadMimeType:  json.upload?.mime_type || 'application/pdf',
        uploadDataB64:   json.upload?.data_b64  || '',
      })
      showToast('Sheet parsed — review and submit when ready', 'success')
    } catch (err) {
      updateDraft(selected.queue_id, { parseStatus: 'error', parseError: err.message || 'parse failed' })
      showToast(err.message || 'Could not parse the upload', 'error')
    }
  }

  const draft = selected ? getDraft(selected) : null
  const effectiveDate = draft?.shiftStartDate || selected?.date || ''

  // Map { employee_name: hours } of hours ALREADY on Daily Sign-In Data
  // for the effectiveDate. The Shift Totals panel adds these to the
  // in-progress crew entries so the user can see whether their worker
  // will end up over the 8h ST cap once this sheet posts. Declared
  // here (NOT later in the function body) because the shiftTotals
  // useMemo a few lines down depends on it — referencing the binding
  // before its `const` would TDZ at component init.
  const [existingDayHours, setExistingDayHours] = useState({})
  // Bump on every successful submit so the existingDayHours useEffect
  // refetches. Without this, navigating from sign-in #1 to sign-in #2
  // in the same shift shows stale totals (sign-in #1's hours don't
  // appear under "Other") until the user reloads the page.
  const [hoursRefreshTick, setHoursRefreshTick] = useState(0)

  // Apply the Mon–Fri-over-8 / Sat-Sun-all-OT rule to a given date+hours.
  const splitStOt = (hours, dateIso) => {
    const m = String(dateIso || '').match(/^(\d{4})-(\d{2})-(\d{2})/)
    const dow = m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getDay() : -1
    const isWeekend = (dow === 0 || dow === 6)
    if (isWeekend) return { st: 0, ot: hours }
    if (hours <= 8) return { st: hours, ot: 0 }
    return { st: 8, ot: hours - 8 }
  }

  // Aggregate { name → { thisSheet, otherSheets, total, st, ot } } for
  // the currently-loaded crew. "thisSheet" = sum of crew rows for that
  // employee in the in-progress form; "otherSheets" = whatever the
  // server already has for them on effectiveDate. Total ST/OT is the
  // shift-wide split, not the per-row split.
  const shiftTotals = useMemo(() => {
    if (!draft) return []
    const sums = new Map()
    draft.crew.forEach(m => {
      const name = (m.name || '').trim()
      if (!name) return
      const h = parseFloat(m.hours) || 0
      sums.set(name, (sums.get(name) || 0) + h)
    })
    // Include any employee with prior hours on this opDay even if they
    // aren't on this sheet, so the user sees the full picture.
    Object.keys(existingDayHours).forEach(n => {
      if (!sums.has(n)) sums.set(n, 0)
    })
    const out = []
    for (const [name, thisSheet] of sums.entries()) {
      const otherSheets = Number(existingDayHours[name] || 0)
      const total = thisSheet + otherSheets
      const split = splitStOt(total, effectiveDate)
      out.push({ name, thisSheet, otherSheets, total, ...split })
    }
    out.sort((a, b) => a.name.localeCompare(b.name))
    return out
  }, [draft, existingDayHours, effectiveDate])

  // Build a Blob URL for the uploaded PDF so react-pdf can render a
  // preview. Revoke when the bytes change so we don't leak.
  const pdfPreviewUrl = useMemo(() => {
    const b64 = draft?.uploadDataB64
    if (!b64) return null
    try {
      const binary = atob(b64)
      const bytes = new Uint8Array(binary.length)
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
      return URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' }))
    } catch (e) {
      return null
    }
  }, [draft?.uploadDataB64])

  useEffect(() => {
    return () => {
      if (pdfPreviewUrl) URL.revokeObjectURL(pdfPreviewUrl)
    }
  }, [pdfPreviewUrl])

  // Editing the shift-start date changes the OT rule (Sat/Sun = all OT).
  // Re-run calcHours on every crew row whenever the date changes so the
  // displayed Hours / OT stay in sync.
  const setShiftStartDate = (newDate) => {
    if (!selected || !draft) return
    const recalculatedCrew = draft.crew.map(m => ({
      ...m,
      ...calcHours(m.timeIn, m.timeOut, newDate),
    }))
    updateDraft(selected.queue_id, {
      shiftStartDate: newDate,
      shiftDateEdited: newDate !== selected.date,
      crew: recalculatedCrew,
    })
  }

  const [editDateModal, setEditDateModal] = useState(false)
  const [dateEditMode,  setDateEditMode]  = useState(false)
  // Set when the server's continuation sanity check returns a likely
  // recent shift this submission should attach to. Holds the previous
  // sign-in's metadata + the cleaned crew list to resume submit with
  // either the previous date (Yes) or effectiveDate (No).
  const [continuationPrompt, setContinuationPrompt] = useState(null)
  // Reset the date-input expansion when switching between queue entries —
  // the modal-acknowledged "yes I'm editing" only applies to the entry
  // that was active at the time of acknowledgement.
  useEffect(() => {
    setDateEditMode(false)
    setEditDateModal(false)
  }, [selectedId])

  // Pull existing Daily Sign-In Data hours for the effectiveDate so the
  // Shift Totals panel can show "this sheet plus what's already logged".
  // Refetch when the date changes (queue switch, kebab edit, etc).
  useEffect(() => {
    if (!effectiveDate) { setExistingDayHours({}); return }
    let cancelled = false
    fetch('/api/signin-queue/day-hours', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ date: effectiveDate }),
    })
      .then(r => r.json())
      .then(json => {
        if (cancelled) return
        setExistingDayHours(json && json.totals ? json.totals : {})
      })
      .catch(err => {
        if (cancelled) return
        console.warn('day-hours fetch failed:', err)
        setExistingDayHours({})
      })
    return () => { cancelled = true }
  }, [effectiveDate, hoursRefreshTick])

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
  const submitWithDate = async (crewClean, dateToSend) => {
    setSubmitting(true)
    try {
      const res = await fetch('/api/signin', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          queue_id:        selected.queue_id,
          date:            dateToSend,
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
          contractor_name:            '',
          contractor_title:           '',
          contractor_signature_b64:   '',
          date_signed:                '',
          source: draft.mode === 'upload' ? 'uploaded' : 'generated',
          upload_blob_b64:            draft.mode === 'upload' ? draft.uploadDataB64 : undefined,
          upload_filename:            draft.mode === 'upload' ? draft.uploadFilename : undefined,
          // Idempotency token. Server caches the response under this id
          // so a retry after a perceived failure returns the cached
          // success rather than writing duplicate Daily Sign-In Data.
          submit_id:                  draft.submitId,
        }),
      })
      const json = await res.json()
      if (!res.ok || json.error) throw new Error(json.error || `HTTP ${res.status}`)

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

      // Refetch existing day-hours so the next entry's Shift Totals
      // panel reflects this submit immediately.
      setHoursRefreshTick(n => n + 1)

      showToast(
        json.duplicate
          ? `Already submitted — no changes`
          : `Sign-in submitted for ${selected.contract_number}`,
        'success'
      )
    } catch (err) {
      showToast(err.message || 'Submit failed', 'error')
    } finally {
      setSubmitting(false)
    }
  }

  const handleSubmit = async () => {
    if (!selected || !draft || submitting) return

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

    // Sanity check: if there's a Daily Sign-In Data row from any
    // contractor whose Time-Out fell within 60 minutes before this
    // submission's Time-In, the user might be filing a continuation
    // of last night's shift. Prompt them to confirm before submit.
    // The user's kebab override (shiftDateEdited=true) skips the check
    // entirely — they've explicitly chosen the date.
    if (!draft.shiftDateEdited) {
      try {
        const checkRes = await fetch('/api/signin-queue/check-continuation', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            time_in:      crewClean[0].timeIn,
            default_date: effectiveDate,
          }),
        })
        const checkJson = await checkRes.json()
        if (checkJson && checkJson.continuation) {
          // Open the prompt; resume on user choice.
          setContinuationPrompt({
            previousDate:     checkJson.previous_date,
            previousContract: checkJson.previous_contract,
            previousTimeOut:  checkJson.previous_time_out,
            gapMinutes:       checkJson.gap_minutes,
            crewClean,
          })
          return
        }
      } catch (err) {
        // Sanity check is best-effort; failures shouldn't block submit.
        console.warn('check-continuation failed; proceeding with default date:', err)
      }
    }

    await submitWithDate(crewClean, effectiveDate)
  }

  // Resolved by the continuation modal. yes = use previous_date; no = use effectiveDate.
  const resolveContinuation = async (useChoice) => {
    if (!continuationPrompt) return
    const { crewClean, previousDate } = continuationPrompt
    const chosen = useChoice === 'yes' ? previousDate : effectiveDate
    setContinuationPrompt(null)
    await submitWithDate(crewClean, chosen)
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
            {/* Header card — read-only context except the kebab-editable date */}
            <div className="card p-4 space-y-2">
              <div className="grid grid-cols-2 gap-3 text-sm">
                <div>
                  <div className="flex items-start justify-between">
                    <span className="field-label">Shift Start Date</span>
                    {!dateEditMode && (
                      <RowKebab items={[{
                        label: 'Edit shift start date',
                        onClick: () => setEditDateModal(true),
                      }]} />
                    )}
                  </div>
                  {dateEditMode ? (
                    <div className="space-y-1">
                      <input
                        type="date"
                        value={effectiveDate}
                        onChange={e => setShiftStartDate(e.target.value)}
                        className="field-input"
                      />
                      <button
                        type="button"
                        onClick={() => setDateEditMode(false)}
                        className="text-[11px] text-navy font-semibold hover:underline">
                        Done editing
                      </button>
                    </div>
                  ) : (
                    <>
                      <p className={`font-semibold ${draft.shiftDateEdited ? 'text-orange-600' : 'text-slate-400'}`}>
                        {prettyQueueDate(effectiveDate)}
                      </p>
                      <p className="text-[11px] text-slate-400">
                        {draft.shiftDateEdited
                          ? 'Edited — overtime applies to this day'
                          : 'Auto-detected from Field Report Date of Work'}
                      </p>
                    </>
                  )}
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

            {/* Mode tabs — switch between filling the form and uploading
                a hand-filled scan that Claude Vision pre-parses. */}
            <div className="card overflow-hidden">
              <div className="grid grid-cols-2">
                {[
                  { key: 'generate', label: 'Generate' },
                  { key: 'upload',   label: 'Upload PDF' },
                ].map(t => (
                  <button
                    key={t.key}
                    type="button"
                    onClick={() => updateDraft(selected.queue_id, { mode: t.key })}
                    className={`py-3 text-sm font-semibold transition-colors
                      ${draft.mode === t.key
                        ? 'bg-navy text-white'
                        : 'bg-white text-slate-500 hover:bg-slate-50'}`}>
                    {t.label}
                  </button>
                ))}
              </div>
            </div>

            {/* Upload zone — visible in upload mode UNTIL parsing completes.
                After parse, the form sections below take over and the
                user reviews/edits prior to submit. */}
            {draft.mode === 'upload' && draft.parseStatus !== 'parsed' && (
              <div className="card p-6 space-y-3">
                <p className="section-label">Upload Hand-Filled Sheet</p>
                {draft.parseStatus === 'parsing' ? (
                  <div className="flex items-center gap-3 py-6 justify-center">
                    <div className="w-5 h-5 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
                    <span className="text-sm text-slate-600">Reading sheet…</span>
                  </div>
                ) : (
                  <>
                    <label className="block border-2 border-dashed border-slate-300 rounded-xl p-8 text-center cursor-pointer hover:border-navy transition-colors">
                      <input
                        type="file"
                        accept="application/pdf"
                        className="hidden"
                        onChange={e => {
                          const f = e.target.files?.[0]
                          if (f) handleUploadFile(f)
                          e.target.value = ''   // allow re-uploading same file
                        }}
                      />
                      <div className="text-3xl mb-1">📄</div>
                      <p className="text-sm font-semibold text-navy">Click to choose PDF</p>
                      <p className="text-[12px] text-slate-500 mt-1">
                        Hand-filled paper sheet, scanned to PDF (≤ 15 MB)
                      </p>
                    </label>
                    {draft.parseStatus === 'error' && (
                      <p className="text-[12px] text-red-600">{draft.parseError}</p>
                    )}
                  </>
                )}
              </div>
            )}

            {/* Once parsed, render the uploaded PDF inline so the user
                can refer back to the handwriting while reviewing the
                parsed values. Stays above the crew rows; user can
                scroll past once they're confident. */}
            {draft.mode === 'upload' && draft.parseStatus === 'parsed' && pdfPreviewUrl && (
              <div className="card p-3 space-y-2">
                <div className="flex items-center justify-between gap-3">
                  <div className="flex items-center gap-2 min-w-0">
                    <span className="text-emerald-700">📎</span>
                    <span className="text-[12px] text-emerald-800 truncate">
                      Uploaded: <span className="font-semibold">{draft.uploadFilename}</span>
                    </span>
                  </div>
                  <button
                    type="button"
                    onClick={() => updateDraft(selected.queue_id, {
                      parseStatus: 'idle', uploadDataB64: '', uploadFilename: '',
                    })}
                    className="text-[11px] text-emerald-700 font-semibold hover:underline flex-shrink-0">
                    Re-upload
                  </button>
                </div>
                <div className="bg-slate-100 border border-slate-200 rounded-lg max-h-[480px] overflow-y-auto p-2 flex justify-center">
                  <UploadPreview url={pdfPreviewUrl} />
                </div>
              </div>
            )}

            {/* Crew + contractor sections — visible in generate mode
                always; in upload mode only AFTER the parse populates
                them. */}
            {(draft.mode === 'generate' || draft.parseStatus === 'parsed') && (
            <>
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
                    workDate={effectiveDate}
                    showSignatures={draft.mode !== 'upload'}
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

            <ShiftTotalsPanel
              totals={shiftTotals}
              dateLabel={prettyQueueDate(effectiveDate)}
            />

            <div className="sticky bottom-4 z-10">
              <button
                onClick={handleSubmit}
                disabled={submitting}
                className="btn-primary w-full text-base">
                {submitting ? 'Submitting…' : 'Submit Sign-In'}
              </button>
            </div>
            </>
            )}
          </div>
        ) : (
          <div className="card p-8 text-center text-slate-500">
            Select an entry from the queue
          </div>
        )}
      </div>

      <Toast message={toast.message} kind={toast.kind} />

      {editDateModal && (
        <ConfirmModal
          title="Edit shift start date?"
          message="The shift start date controls overtime calculation (Sat/Sun = all OT, weekday over 8 = OT). Only edit this if the work actually started on a different day than the Field Report."
          confirmLabel="Edit anyway"
          cancelLabel="Cancel"
          onConfirm={() => { setEditDateModal(false); setDateEditMode(true) }}
          onCancel={() => setEditDateModal(false)}
        />
      )}

      {continuationPrompt && (
        <ConfirmModal
          title="Part of last night's shift?"
          message={`A recent sign-in for ${continuationPrompt.previousContract} signed out at ${continuationPrompt.previousTimeOut} on ${prettyQueueDate(continuationPrompt.previousDate)} — only ${continuationPrompt.gapMinutes} min before this Time In. Bucket this sign-in under ${prettyQueueDate(continuationPrompt.previousDate)} so OT and certified payroll attribute the hours to that shift?`}
          confirmLabel={`Yes — bucket under ${prettyQueueDate(continuationPrompt.previousDate)}`}
          cancelLabel="No — separate shift"
          onConfirm={() => resolveContinuation('yes')}
          onCancel={() => resolveContinuation('no')}
        />
      )}
    </div>
  )
}
