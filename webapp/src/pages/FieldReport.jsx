import { useState, useEffect, useCallback } from 'react'
import StatusBadge from '../components/StatusBadge'

// ── Constants ─────────────────────────────────────────────────
const CLASSIFICATIONS = [
  'LP (Crew Chief)',
  'Laborer',
  'Driver',
  'Flagger',
  'Foreman',
  'Equipment Operator'
]

const MARKING_TYPES = [
  'Crosswalk',
  'Stop Bar',
  'Lane Line',
  'Center Line',
  'Double Yellow',
  'No Passing Zone',
  'Arrow – Straight',
  'Arrow – Left',
  'Arrow – Right',
  'Arrow – U-Turn',
  'Combined Arrow',
  'Bike Lane',
  'Bus Lane',
  'Pedestrian Space',
  'Yield Line',
  'Yield to Pedestrians',
  'School (SCHOOL)',
  'Railroad Crossing',
  'Parking Space',
  'Curb Marking',
  'Other'
]

// ── Helpers ───────────────────────────────────────────────────
function isoToday() {
  const d = new Date()
  return [
    d.getFullYear(),
    String(d.getMonth() + 1).padStart(2, '0'),
    String(d.getDate()).padStart(2, '0')
  ].join('-')
}

function fmt24to12(t) {
  if (!t) return ''
  const [h, m] = t.split(':').map(Number)
  const ampm = h >= 12 ? 'PM' : 'AM'
  return (h % 12 || 12) + ':' + String(m).padStart(2, '0') + ' ' + ampm
}

function calcHoursFromTimes(timeIn, timeOut) {
  if (!timeIn || !timeOut) return { hours: '', overtime: '' }
  const [ih, im] = timeIn.split(':').map(Number)
  const [oh, om] = timeOut.split(':').map(Number)
  const mins = (oh * 60 + om) - (ih * 60 + im)
  if (mins <= 0) return { hours: '', overtime: '' }
  const hrs = mins / 60
  const ot  = Math.max(0, hrs - 8)
  return { hours: hrs.toFixed(2), overtime: ot.toFixed(2) }
}

function toDateInput(s) {
  if (!s || s === 'null' || s === 'undefined') return ''
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)
  if (m) return `${m[3]}-${m[1].padStart(2,'0')}-${m[2].padStart(2,'0')}`
  return ''
}

// ── Sub-components ────────────────────────────────────────────

function Field({ label, required, children, hint }) {
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

function YNToggle({ value, onChange, yesLabel = 'Yes', noLabel = 'No' }) {
  return (
    <div className="flex border border-slate-200 rounded-lg overflow-hidden">
      <button
        type="button"
        onClick={() => onChange('no')}
        className={`flex-1 py-2.5 text-sm font-semibold transition-all
          ${value === 'no'
            ? 'bg-red-500 text-white'
            : 'bg-white text-slate-500 hover:bg-slate-50'}`}
      >
        {noLabel}
      </button>
      <button
        type="button"
        onClick={() => onChange('yes')}
        className={`flex-1 py-2.5 text-sm font-semibold border-l border-slate-200 transition-all
          ${value === 'yes'
            ? 'bg-green-600 text-white'
            : 'bg-white text-slate-500 hover:bg-slate-50'}`}
      >
        {yesLabel}
      </button>
    </div>
  )
}

function MarkingRow({ idx, data, onChange, onRemove }) {
  return (
    <div className="grid grid-cols-[1fr_72px_60px_32px] gap-1.5 items-center">
      <select
        value={data.type}
        onChange={e => onChange(idx, 'type', e.target.value)}
        className="field-input text-sm py-2"
      >
        {MARKING_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
      </select>
      <input
        type="number"
        min="0"
        placeholder="Qty"
        value={data.qty}
        onChange={e => onChange(idx, 'qty', e.target.value)}
        className="field-input text-sm py-2 text-center"
      />
      <select
        value={data.unit}
        onChange={e => onChange(idx, 'unit', e.target.value)}
        className="field-input text-sm py-2"
      >
        <option value="SF">SF</option>
        <option value="LF">LF</option>
        <option value="EA">EA</option>
      </select>
      <button
        type="button"
        onClick={() => onRemove(idx)}
        className="text-red-400 hover:text-red-600 text-xl font-light leading-none
                   transition-colors flex items-center justify-center h-9"
      >
        ×
      </button>
    </div>
  )
}

function CrewCard({ idx, data, onChange, onRemove }) {
  const { hours, overtime } = calcHoursFromTimes(data.timeIn, data.timeOut)

  // Update derived hours whenever time fields change
  const handleTime = (field, val) => {
    const updated = { ...data, [field]: val }
    const derived = calcHoursFromTimes(
      field === 'timeIn'  ? val : data.timeIn,
      field === 'timeOut' ? val : data.timeOut
    )
    onChange(idx, { ...updated, hours: derived.hours, overtime: derived.overtime })
  }

  return (
    <div className="border border-slate-200 rounded-xl p-4 space-y-3 bg-slate-50/50">
      <div className="flex justify-between items-center">
        <span className="text-xs font-bold text-navy uppercase tracking-wider">
          Crew Member #{idx + 1}
        </span>
        <button
          type="button"
          onClick={() => onRemove(idx)}
          className="text-red-400 hover:text-red-600 text-xl leading-none transition-colors"
        >
          ×
        </button>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <Field label="Full Name" required>
          <input
            type="text"
            value={data.name}
            onChange={e => onChange(idx, { ...data, name: e.target.value })}
            placeholder="First Last"
            autoCapitalize="words"
            className="field-input"
          />
        </Field>
        <Field label="Classification" required>
          <select
            value={data.classification}
            onChange={e => onChange(idx, { ...data, classification: e.target.value })}
            className="field-input"
          >
            {CLASSIFICATIONS.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
        </Field>
      </div>

      <div className="grid grid-cols-2 gap-3">
        <Field label="Time In" required>
          <input
            type="time"
            value={data.timeIn}
            onChange={e => handleTime('timeIn', e.target.value)}
            className="field-input"
          />
        </Field>
        <Field label="Time Out" required>
          <input
            type="time"
            value={data.timeOut}
            onChange={e => handleTime('timeOut', e.target.value)}
            className="field-input"
          />
        </Field>
      </div>

      {/* Computed hours display */}
      {data.hours && (
        <div className="flex gap-4 bg-white rounded-lg px-3 py-2 border border-slate-200">
          <div>
            <span className="text-[10px] text-slate-400 uppercase tracking-wider font-bold">Reg. Hours</span>
            <p className="text-sm font-bold text-navy">{data.hours}</p>
          </div>
          <div className="border-l border-slate-200 pl-4">
            <span className="text-[10px] text-slate-400 uppercase tracking-wider font-bold">OT Hours</span>
            <p className={`text-sm font-bold ${parseFloat(data.overtime) > 0 ? 'text-orange-600' : 'text-slate-400'}`}>
              {data.overtime || '0.00'}
            </p>
          </div>
        </div>
      )}
    </div>
  )
}

// ── WO Info Panel ─────────────────────────────────────────────
function WOPanel({ wo }) {
  if (!wo) return null
  const rows = [
    ['Contractor', wo.contractor],
    ['Borough',    wo.borough],
    ['Location',   wo.location],
    ['From → To',  [wo.from_street, wo.to_street].filter(Boolean).join(' → ') || '—'],
    ['Due Date',   wo.due_date || '—'],
    ['Work Type',  wo.work_type || '—'],
  ]
  return (
    <div className="bg-slate-50 rounded-xl p-3 border border-slate-200 space-y-1.5">
      <div className="flex items-center justify-between mb-1">
        <span className="text-[10px] font-extrabold uppercase tracking-wider text-slate-400">
          WO Details
        </span>
        <StatusBadge status={wo.status} />
      </div>
      {rows.map(([key, val]) => (
        <div key={key} className="flex gap-3">
          <span className="text-[11px] text-slate-400 font-semibold min-w-[72px]">{key}</span>
          <span className="text-sm text-slate-700 font-medium">{val}</span>
        </div>
      ))}
    </div>
  )
}

// ── Main Page ─────────────────────────────────────────────────
function newCrew() {
  return { name: '', classification: CLASSIFICATIONS[0], timeIn: '', timeOut: '', hours: '', overtime: '' }
}
function newMarking() {
  return { type: MARKING_TYPES[0], qty: '', unit: 'SF' }
}

export default function FieldReport() {
  // ── Data loading
  const [wos,     setWOs]     = useState([])
  const [loading, setLoading] = useState(true)
  const [apiError,setApiError]= useState(null)

  // ── Form state
  const [selectedWOId,   setSelectedWOId]   = useState('')
  const [workDate,       setWorkDate]        = useState(isoToday())
  const [dispatchDate,   setDispatchDate]    = useState('')
  const [workStartDate,  setWorkStartDate]   = useState('')
  const [workEndDate,    setWorkEndDate]     = useState('')
  const [markingRows,    setMarkingRows]     = useState([newMarking()])
  const [sqft,           setSqft]           = useState('')
  const [paintMaterial,  setPaintMaterial]  = useState('')
  const [issues,         setIssues]         = useState('')
  const [crewMembers,    setCrewMembers]     = useState([newCrew()])
  const [woComplete,     setWoComplete]      = useState('no')
  const [photosUploaded, setPhotosUploaded]  = useState('no')

  // ── UI state
  const [formError,  setFormError]  = useState('')
  const [submitting, setSubmitting] = useState(false)
  const [submitted,  setSubmitted]  = useState(null)   // { wo_id, status, crew_count }

  const selectedWO = wos.find(w => w.id === selectedWOId) ?? null

  // Load active WOs on mount
  useEffect(() => {
    fetch('/api/wos')
      .then(r => r.json())
      .then(d => {
        if (d.error) throw new Error(d.error)
        setWOs(d.wos ?? [])
      })
      .catch(e => setApiError(e.message))
      .finally(() => setLoading(false))
  }, [])

  // Pre-fill date fields when WO is selected
  useEffect(() => {
    if (!selectedWO) return
    setDispatchDate(toDateInput(selectedWO.dispatch_date))
    setWorkStartDate(toDateInput(selectedWO.work_start_date))
  }, [selectedWOId])

  // Show work-end date field when WO is marked complete
  useEffect(() => {
    if (woComplete === 'yes' && !workEndDate) setWorkEndDate(isoToday())
  }, [woComplete])

  // ── Marking row handlers
  const updateMarking = (idx, field, val) =>
    setMarkingRows(rows => rows.map((r, i) => i === idx ? { ...r, [field]: val } : r))
  const addMarking    = () => setMarkingRows(r => [...r, newMarking()])
  const removeMarking = (idx) => setMarkingRows(r => r.filter((_, i) => i !== idx))

  // ── Crew handlers
  const updateCrew = (idx, updated) =>
    setCrewMembers(list => list.map((m, i) => i === idx ? updated : m))
  const addCrew    = () => setCrewMembers(list => [...list, newCrew()])
  const removeCrew = (idx) => setCrewMembers(list => list.filter((_, i) => i !== idx))

  // ── Build marking types string
  const markingTypesStr = markingRows
    .filter(r => r.qty && parseFloat(r.qty) > 0)
    .map(r => `${r.type}: ${r.qty} ${r.unit}`)
    .join(', ')

  // ── Submit
  async function handleSubmit(e) {
    e.preventDefault()
    setFormError('')

    if (!selectedWOId)     { setFormError('Please select a Work Order.'); return }
    if (!workDate)         { setFormError('Please enter the date of work.'); return }
    const validCrew = crewMembers.filter(m => m.name.trim())
    if (validCrew.length === 0) { setFormError('Please add at least one crew member.'); return }
    for (const m of validCrew) {
      if (!m.timeIn || !m.timeOut) {
        setFormError(`Please enter Time In and Time Out for ${m.name || 'all crew members'}.`)
        return
      }
    }

    setSubmitting(true)
    try {
      const payload = {
        wo_id:           selectedWOId,
        date:            workDate,
        dispatch_date:   dispatchDate,
        work_start_date: workStartDate,
        work_end_date:   workEndDate,
        wo_complete:     woComplete === 'yes',
        marking_types:   markingTypesStr,
        sqft_completed:  sqft !== '' ? parseFloat(sqft) : null,
        paint_material:  paintMaterial.trim(),
        issues:          issues.trim(),
        photos_uploaded: photosUploaded === 'yes',
        crew: validCrew.map(m => ({
          name:           m.name.trim(),
          classification: m.classification,
          time_in:        fmt24to12(m.timeIn),
          time_out:       fmt24to12(m.timeOut),
          hours:          parseFloat(m.hours)   || 0,
          overtime:       parseFloat(m.overtime) || 0
        }))
      }

      const res  = await fetch('/api/field-report', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify(payload)
      })
      const data = await res.json()
      if (data.error) throw new Error(data.error)

      setSubmitted({ wo_id: data.wo_id, status: data.status, crew_count: validCrew.length })
    } catch (err) {
      setFormError('Submission failed: ' + err.message)
    } finally {
      setSubmitting(false)
    }
  }

  // ── Success screen
  if (submitted) {
    return (
      <div className="max-w-lg mx-auto px-4 py-16 text-center">
        <div className="text-6xl mb-6">✅</div>
        <h2 className="text-2xl font-black text-navy mb-3">Report Submitted!</h2>
        <p className="text-slate-500 mb-2">
          <strong className="text-navy font-mono">{submitted.wo_id}</strong>
          {' '}— {submitted.crew_count} crew member(s) logged
        </p>
        <div className="mb-8">
          <StatusBadge status={submitted.status} />
        </div>
        <div className="flex gap-3 justify-center flex-wrap">
          <button
            onClick={() => {
              setSubmitted(null)
              setSelectedWOId('')
              setWorkDate(isoToday())
              setDispatchDate(''); setWorkStartDate(''); setWorkEndDate('')
              setMarkingRows([newMarking()])
              setSqft(''); setPaintMaterial(''); setIssues('')
              setCrewMembers([newCrew()])
              setWoComplete('no'); setPhotosUploaded('no')
            }}
            className="btn-primary px-6"
          >
            Submit Another Report
          </button>
          <a href="/" className="btn-outline px-6">Go to Dashboard</a>
        </div>
      </div>
    )
  }

  // ── Loading / error
  if (loading) {
    return (
      <div className="flex items-center justify-center min-h-[50vh]">
        <div className="flex flex-col items-center gap-4">
          <div className="w-9 h-9 border-[3px] border-slate-200 border-t-navy rounded-full animate-spin" />
          <p className="text-slate-400 text-sm">Loading work orders…</p>
        </div>
      </div>
    )
  }

  if (apiError) {
    return (
      <div className="flex items-center justify-center min-h-[50vh]">
        <div className="text-center max-w-sm">
          <div className="text-4xl mb-3">⚠️</div>
          <p className="text-slate-500 text-sm mb-4">{apiError}</p>
          <button onClick={() => window.location.reload()} className="btn-outline text-sm px-4 py-2">
            Reload
          </button>
        </div>
      </div>
    )
  }

  // ── Main form
  return (
    <div className="max-w-2xl mx-auto px-4 py-6 space-y-4">

      <div>
        <h1 className="text-2xl font-black text-navy">Field Report</h1>
        <p className="text-slate-500 text-sm mt-0.5">
          Log today's crew sign-in and work progress for a Work Order.
        </p>
      </div>

      {/* Global form error */}
      {formError && (
        <div className="bg-red-50 border border-red-200 text-red-700 rounded-xl p-4 text-sm">
          {formError}
        </div>
      )}

      <form onSubmit={handleSubmit} className="space-y-4">

        {/* ── 1. Work Order ──────────────────────────── */}
        <div className="card p-4 space-y-3">
          <p className="section-label">Work Order</p>

          <Field label="Select Work Order" required>
            <select
              value={selectedWOId}
              onChange={e => setSelectedWOId(e.target.value)}
              className="field-input"
            >
              <option value="">— Choose a Work Order —</option>
              {wos.map(wo => (
                <option key={wo.id} value={wo.id}>
                  {wo.id} — {wo.location} ({wo.borough})
                </option>
              ))}
            </select>
          </Field>

          {selectedWO && <WOPanel wo={selectedWO} />}
        </div>

        {/* ── 2. Schedule ───────────────────────────── */}
        <div className="card p-4 space-y-3">
          <p className="section-label">Schedule</p>

          <div className="grid grid-cols-2 gap-3">
            <Field label="Date of Work" required>
              <input
                type="date"
                value={workDate}
                onChange={e => setWorkDate(e.target.value)}
                className="field-input"
              />
            </Field>
            <Field
              label="Dispatch Date"
              hint={selectedWO?.dispatch_date ? `Already set: ${selectedWO.dispatch_date}` : undefined}
            >
              <input
                type="date"
                value={dispatchDate}
                onChange={e => setDispatchDate(e.target.value)}
                className="field-input"
              />
            </Field>
          </div>

          <div className="grid grid-cols-2 gap-3">
            <Field
              label="Work Start Date"
              hint={selectedWO?.work_start_date ? `Already set: ${selectedWO.work_start_date}` : undefined}
            >
              <input
                type="date"
                value={workStartDate}
                onChange={e => setWorkStartDate(e.target.value)}
                className="field-input"
              />
            </Field>
            {woComplete === 'yes' && (
              <Field label="Work End Date">
                <input
                  type="date"
                  value={workEndDate}
                  onChange={e => setWorkEndDate(e.target.value)}
                  className="field-input"
                />
              </Field>
            )}
          </div>

          <p className="text-[11px] text-slate-400">
            Only fill Dispatch / Start / End dates if not already in the tracker.
          </p>
        </div>

        {/* ── 3. Work Details ───────────────────────── */}
        <div className="card p-4 space-y-4">
          <p className="section-label">Work Details</p>

          <div className="space-y-2">
            <label className="field-label">Marking Types Completed</label>
            <div className="space-y-1.5">
              {markingRows.map((row, idx) => (
                <MarkingRow
                  key={idx}
                  idx={idx}
                  data={row}
                  onChange={updateMarking}
                  onRemove={removeMarking}
                />
              ))}
            </div>
            <button type="button" onClick={addMarking} className="btn-ghost text-xs">
              + Add Marking Type
            </button>
            {markingTypesStr && (
              <p className="text-[11px] text-slate-400 bg-slate-50 rounded-lg px-2 py-1.5 font-mono">
                {markingTypesStr}
              </p>
            )}
          </div>

          <div className="grid grid-cols-2 gap-3">
            <Field label="SQFT Completed">
              <input
                type="number"
                value={sqft}
                onChange={e => setSqft(e.target.value)}
                placeholder="0"
                min="0"
                inputMode="numeric"
                className="field-input"
              />
            </Field>
            <Field label="Paint / Material">
              <input
                type="text"
                value={paintMaterial}
                onChange={e => setPaintMaterial(e.target.value)}
                placeholder="e.g. White Thermo"
                className="field-input"
              />
            </Field>
          </div>

          <Field label="Issues / Notes">
            <textarea
              value={issues}
              onChange={e => setIssues(e.target.value)}
              placeholder="Problems, delays, or anything admin should know…"
              rows={3}
              className="field-input resize-none"
            />
          </Field>
        </div>

        {/* ── 4. Crew ───────────────────────────────── */}
        <div className="card p-4 space-y-3">
          <p className="section-label">Crew ({crewMembers.length})</p>

          <div className="space-y-3">
            {crewMembers.map((member, idx) => (
              <CrewCard
                key={idx}
                idx={idx}
                data={member}
                onChange={updateCrew}
                onRemove={removeCrew}
              />
            ))}
          </div>

          <button type="button" onClick={addCrew} className="btn-ghost text-xs w-full">
            + Add Crew Member
          </button>
        </div>

        {/* ── 5. Completion ─────────────────────────── */}
        <div className="card p-4 space-y-4">
          <p className="section-label">Completion</p>

          <Field label="Is this Work Order complete?">
            <YNToggle
              value={woComplete}
              onChange={setWoComplete}
              noLabel="No — more work needed"
              yesLabel="Yes — WO Done ✓"
            />
          </Field>

          <Field label="Photos uploaded to Drive?">
            <YNToggle
              value={photosUploaded}
              onChange={setPhotosUploaded}
              noLabel="Not yet"
              yesLabel="Yes — uploaded ✓"
            />
          </Field>
        </div>

        {/* ── Submit ────────────────────────────────── */}
        <button
          type="submit"
          disabled={submitting}
          className="btn-primary w-full text-base"
        >
          {submitting ? 'Submitting…' : 'Submit Field Report'}
        </button>

        <div className="h-8" />
      </form>
    </div>
  )
}
