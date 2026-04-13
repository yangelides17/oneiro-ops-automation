import { useState, useEffect, useRef } from 'react'
import StatusBadge  from '../components/StatusBadge'
import SignaturePad from '../components/SignaturePad'
import ConfirmModal from '../components/ConfirmModal'

// ── Constants ─────────────────────────────────────────────────
const CLASSIFICATIONS = [
  'LP (Crew Chief)', 'Laborer', 'Driver', 'Flagger', 'Foreman', 'Equipment Operator'
]
const MARKING_TYPES = [
  'Crosswalk', 'Stop Bar', 'Lane Line', 'Center Line', 'Double Yellow',
  'No Passing Zone', 'Arrow – Straight', 'Arrow – Left', 'Arrow – Right',
  'Arrow – U-Turn', 'Combined Arrow', 'Bike Lane', 'Bus Lane',
  'Pedestrian Space', 'Yield Line', 'Yield to Pedestrians', 'School (SCHOOL)',
  'Railroad Crossing', 'Parking Space', 'Curb Marking', 'Other'
]

// ── Helpers ───────────────────────────────────────────────────
const isoToday = () => {
  const d = new Date()
  return [d.getFullYear(), String(d.getMonth()+1).padStart(2,'0'), String(d.getDate()).padStart(2,'0')].join('-')
}
const fmt24to12 = (t) => {
  if (!t) return ''
  const [h, m] = t.split(':').map(Number)
  return (h%12||12) + ':' + String(m).padStart(2,'0') + (h>=12?' PM':' AM')
}
const calcHours = (tin, tout) => {
  if (!tin || !tout) return { hours: '', overtime: '' }
  const [ih,im] = tin.split(':').map(Number)
  const [oh,om] = tout.split(':').map(Number)
  const mins = (oh*60+om) - (ih*60+im)
  if (mins <= 0) return { hours: '', overtime: '' }
  const hrs = mins / 60
  return { hours: hrs.toFixed(2), overtime: Math.max(0, hrs-8).toFixed(2) }
}
const newCrew    = () => ({ name:'', classification:CLASSIFICATIONS[0], timeIn:'', timeOut:'', hours:'', overtime:'', signatureIn:null, signatureOut:null })
const newMarking = () => ({ type:MARKING_TYPES[0], qty:'', unit:'SF' })

// ── Field wrapper ─────────────────────────────────────────────
function Field({ label, required, hint, children }) {
  return (
    <div className="space-y-1">
      <label className="field-label">{label}{required && <span className="text-red-400 ml-0.5">*</span>}</label>
      {children}
      {hint && <p className="text-[11px] text-slate-400">{hint}</p>}
    </div>
  )
}

// ── Yes / No toggle ───────────────────────────────────────────
function YNToggle({ value, onChange, yesLabel='Yes', noLabel='No' }) {
  return (
    <div className="flex border border-slate-200 rounded-lg overflow-hidden">
      {[{v:'no',l:noLabel,cls:'bg-red-500 text-white'},{v:'yes',l:yesLabel,cls:'bg-green-600 text-white'}].map(({v,l,cls})=>(
        <button key={v} type="button" onClick={()=>onChange(v)}
          className={`flex-1 py-2.5 text-sm font-semibold border-l first:border-l-0 border-slate-200 transition-all
            ${value===v ? cls : 'bg-white text-slate-500 hover:bg-slate-50'}`}>
          {l}
        </button>
      ))}
    </div>
  )
}

// ── Marking row ───────────────────────────────────────────────
function MarkingRow({ idx, data, onChange, onRemove }) {
  return (
    <div className="grid grid-cols-[1fr_72px_60px_32px] gap-1.5 items-center">
      <select value={data.type} onChange={e=>onChange(idx,'type',e.target.value)} className="field-input text-sm py-2">
        {MARKING_TYPES.map(t=><option key={t}>{t}</option>)}
      </select>
      <input type="number" min="0" placeholder="Qty" value={data.qty}
        onChange={e=>onChange(idx,'qty',e.target.value)} className="field-input text-sm py-2 text-center" />
      <select value={data.unit} onChange={e=>onChange(idx,'unit',e.target.value)} className="field-input text-sm py-2">
        {['SF','LF','EA'].map(u=><option key={u}>{u}</option>)}
      </select>
      <button type="button" onClick={()=>onRemove(idx)}
        className="text-red-400 hover:text-red-600 text-xl font-light h-9 flex items-center justify-center">×</button>
    </div>
  )
}

// ── Crew card (with signature pads) ──────────────────────────
function CrewCard({ idx, data, onChange, onRemove }) {
  const handleTime = (field, val) => {
    const tin  = field==='timeIn'  ? val : data.timeIn
    const tout = field==='timeOut' ? val : data.timeOut
    onChange(idx, { ...data, [field]: val, ...calcHours(tin, tout) })
  }
  return (
    <div className="border border-slate-200 rounded-xl p-4 space-y-4 bg-slate-50/40">
      <div className="flex justify-between items-center">
        <span className="text-xs font-bold text-navy uppercase tracking-wider">Crew Member #{idx+1}</span>
        <button type="button" onClick={()=>onRemove(idx)} className="text-red-400 hover:text-red-600 text-xl leading-none">×</button>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <Field label="Full Name" required>
          <input type="text" value={data.name} autoCapitalize="words" placeholder="First Last"
            onChange={e=>onChange(idx,{...data,name:e.target.value})} className="field-input" />
        </Field>
        <Field label="Classification" required>
          <select value={data.classification} onChange={e=>onChange(idx,{...data,classification:e.target.value})} className="field-input">
            {CLASSIFICATIONS.map(c=><option key={c}>{c}</option>)}
          </select>
        </Field>
      </div>

      <div className="grid grid-cols-2 gap-3">
        <Field label="Time In" required>
          <input type="time" value={data.timeIn} onChange={e=>handleTime('timeIn',e.target.value)} className="field-input" />
        </Field>
        <Field label="Time Out" required>
          <input type="time" value={data.timeOut} onChange={e=>handleTime('timeOut',e.target.value)} className="field-input" />
        </Field>
      </div>

      {data.hours && (
        <div className="flex gap-4 bg-white rounded-lg px-3 py-2 border border-slate-200">
          <div>
            <span className="text-[10px] text-slate-400 uppercase tracking-wider font-bold">Reg. Hours</span>
            <p className="text-sm font-bold text-navy">{data.hours}</p>
          </div>
          <div className="border-l border-slate-200 pl-4">
            <span className="text-[10px] text-slate-400 uppercase tracking-wider font-bold">OT Hours</span>
            <p className={`text-sm font-bold ${parseFloat(data.overtime)>0?'text-orange-600':'text-slate-400'}`}>
              {data.overtime||'0.00'}
            </p>
          </div>
        </div>
      )}

      {/* Signature pads — captured locally, added to Sign-In PDF later */}
      <div className="pt-1 border-t border-slate-200 space-y-3">
        <p className="text-[11px] text-slate-400 font-medium">
          Employee signs below — signatures will be added to the Sign-In Log
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

// ── WO info panel ─────────────────────────────────────────────
function WOPanel({ wo }) {
  return (
    <div className="bg-slate-50 rounded-xl p-3 border border-slate-200 space-y-1.5">
      <div className="flex items-center justify-between mb-1">
        <span className="text-[10px] font-extrabold uppercase tracking-wider text-slate-400">WO Details</span>
        <StatusBadge status={wo.status} />
      </div>
      {[
        ['Contractor', wo.contractor],
        ['Borough',    wo.borough],
        ['Location',   wo.location],
        ['From → To',  [wo.from_street,wo.to_street].filter(Boolean).join(' → ')||'—'],
        ['Due Date',   wo.due_date||'—'],
        ['Work Type',  wo.work_type||'—'],
      ].map(([k,v])=>(
        <div key={k} className="flex gap-3">
          <span className="text-[11px] text-slate-400 font-semibold min-w-[72px]">{k}</span>
          <span className="text-sm text-slate-700 font-medium">{v}</span>
        </div>
      ))}
    </div>
  )
}

// ── Photo picker (upload happens on submit) ───────────────────
function PhotoPicker({ onChange }) {
  const fileRef = useRef(null)
  const [files, setFiles] = useState([])

  const addFiles = (e) => {
    const added   = Array.from(e.target.files || [])
    const updated = [...files, ...added]
    setFiles(updated)
    onChange(updated)
    e.target.value = ''
  }
  const remove = (idx) => {
    const updated = files.filter((_,i)=>i!==idx)
    setFiles(updated)
    onChange(updated)
  }

  return (
    <div className="space-y-3">
      {/* Drop zone */}
      <div onClick={()=>fileRef.current?.click()}
        className="border-2 border-dashed border-slate-200 rounded-xl p-5 text-center cursor-pointer
                   hover:border-navy/40 hover:bg-slate-50 transition-all select-none">
        <p className="text-2xl mb-1">📷</p>
        <p className="text-sm font-semibold text-slate-600">Tap to select photos</p>
        <p className="text-xs text-slate-400 mt-0.5">JPEG, PNG, HEIC — up to 15 MB each</p>
        <input ref={fileRef} type="file" multiple accept="image/*" className="hidden" onChange={addFiles} />
      </div>

      {/* Thumbnails */}
      {files.length > 0 && (
        <>
          <div className="flex flex-wrap gap-2">
            {files.map((f,i)=>(
              <div key={i} className="relative group">
                <img src={URL.createObjectURL(f)} alt={f.name}
                  className="w-16 h-16 object-cover rounded-lg border border-slate-200" />
                <button type="button" onClick={()=>remove(i)}
                  className="absolute -top-1.5 -right-1.5 w-5 h-5 bg-red-500 text-white rounded-full
                             text-xs font-bold flex items-center justify-center
                             opacity-0 group-hover:opacity-100 transition-opacity">×</button>
              </div>
            ))}
          </div>
          <p className="text-[11px] text-green-700 font-semibold">
            ✓ {files.length} photo{files.length!==1?'s':''} selected — will upload to Drive on submit
          </p>
        </>
      )}

      {files.length === 0 && (
        <p className="text-[11px] text-slate-400">
          Photos upload to Drive automatically when you submit the report.
        </p>
      )}
    </div>
  )
}

// ── Main page ─────────────────────────────────────────────────
export default function FieldReport() {
  // Data
  const [wos,      setWOs]      = useState([])
  const [loading,  setLoading]  = useState(true)
  const [apiError, setApiError] = useState(null)

  // Form fields
  const [selectedWOId,  setSelectedWOId]  = useState('')
  const [workDate,      setWorkDate]       = useState(isoToday())
  const [markingRows,   setMarkingRows]    = useState([newMarking()])
  const [sqft,          setSqft]           = useState('')
  const [paintMaterial, setPaintMaterial]  = useState('')
  const [issues,        setIssues]         = useState('')
  const [crewMembers,   setCrewMembers]    = useState([newCrew()])
  const [woComplete,    setWoComplete]     = useState('no')
  const [photoFiles,    setPhotoFiles]     = useState([])   // File objects, uploaded on submit

  // UI state
  const [formError,      setFormError]      = useState('')
  const [submitStep,     setSubmitStep]     = useState('')  // current step label
  const [submitting,     setSubmitting]     = useState(false)
  const [submitted,      setSubmitted]      = useState(null)
  const [showPhotoModal, setShowPhotoModal] = useState(false)

  const selectedWO = wos.find(w => w.id === selectedWOId) ?? null

  useEffect(() => {
    fetch('/api/wos').then(r=>r.json())
      .then(d=>{ if(d.error) throw new Error(d.error); setWOs(d.wos??[]) })
      .catch(e=>setApiError(e.message))
      .finally(()=>setLoading(false))
  }, [])

  // Marking rows
  const updateMarking = (i,f,v) => setMarkingRows(r=>r.map((x,j)=>j===i?{...x,[f]:v}:x))
  const addMarking    = () => setMarkingRows(r=>[...r,newMarking()])
  const removeMarking = (i) => setMarkingRows(r=>r.filter((_,j)=>j!==i))

  // Crew rows
  const updateCrew = (i, updated) => setCrewMembers(l=>l.map((m,j)=>j===i?updated:m))
  const addCrew    = () => setCrewMembers(l=>[...l,newCrew()])
  const removeCrew = (i) => setCrewMembers(l=>l.filter((_,j)=>j!==i))

  const markingTypesStr = markingRows
    .filter(r=>r.qty && parseFloat(r.qty)>0)
    .map(r=>`${r.type}: ${r.qty} ${r.unit}`).join(', ')

  // ── Core submit (called after all guards pass) ────────────
  async function doSubmit() {
    setFormError('')
    setSubmitting(true)
    const validCrew = crewMembers.filter(m=>m.name.trim())

    // Step 1 — upload photos to Drive
    let photosUploaded = false
    if (photoFiles.length > 0) {
      for (let i = 0; i < photoFiles.length; i++) {
        setSubmitStep(`Uploading photo ${i+1} of ${photoFiles.length}…`)
        const form = new FormData()
        form.append('photo', photoFiles[i])
        form.append('wo_id', selectedWOId)
        try {
          const res  = await fetch('/api/upload-photo', { method:'POST', body:form })
          const data = await res.json()
          if (!data.error) photosUploaded = true
        } catch { /* non-fatal */ }
      }
    }

    // Step 2 — submit field report
    setSubmitStep('Submitting report…')
    try {
      const res = await fetch('/api/field-report', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          wo_id:       selectedWOId,
          date:        workDate,
          wo_complete: woComplete==='yes',
          marking_types:   markingTypesStr,
          sqft_completed:  sqft!=='' ? parseFloat(sqft) : null,
          paint_material:  paintMaterial.trim(),
          issues:          issues.trim(),
          photos_uploaded: photosUploaded,
          crew: validCrew.map(m=>({
            name:           m.name.trim(),
            classification: m.classification,
            time_in:        fmt24to12(m.timeIn),
            time_out:       fmt24to12(m.timeOut),
            hours:          parseFloat(m.hours)||0,
            overtime:       parseFloat(m.overtime)||0,
            signed_in:      !!m.signatureIn,
            signed_out:     !!m.signatureOut
          }))
        })
      })
      const data = await res.json()
      if (data.error) throw new Error(data.error)
      setSubmitted({ wo_id:data.wo_id, status:data.status, crew_count:validCrew.length, photos:photoFiles.length })
    } catch (err) {
      setFormError('Submission failed: ' + err.message)
    } finally {
      setSubmitting(false)
      setSubmitStep('')
    }
  }

  // ── Form submit — validate then guard → doSubmit ──────────
  function handleSubmit(e) {
    e.preventDefault()
    setFormError('')
    if (!selectedWOId)  { setFormError('Please select a Work Order.'); return }
    if (!workDate)      { setFormError('Please enter the date of work.'); return }
    const valid = crewMembers.filter(m=>m.name.trim())
    if (!valid.length)  { setFormError('Please add at least one crew member.'); return }
    for (const m of valid) {
      if (!m.timeIn||!m.timeOut) {
        setFormError(`Enter Time In and Time Out for ${m.name||'all crew members'}.`); return
      }
    }
    // Guard: no photos selected
    if (photoFiles.length===0) { setShowPhotoModal(true); return }
    doSubmit()
  }

  // ── Reset ─────────────────────────────────────────────────
  function reset() {
    setSubmitted(null); setSelectedWOId(''); setWorkDate(isoToday())
    setMarkingRows([newMarking()]); setSqft(''); setPaintMaterial(''); setIssues('')
    setCrewMembers([newCrew()]); setWoComplete('no'); setPhotoFiles([]); setFormError('')
  }

  // ── Success ───────────────────────────────────────────────
  if (submitted) return (
    <div className="max-w-lg mx-auto px-4 py-16 text-center">
      <div className="text-6xl mb-5">✅</div>
      <h2 className="text-2xl font-black text-navy mb-2">Report Submitted!</h2>
      <p className="text-slate-500 mb-1">
        <strong className="font-mono text-navy">{submitted.wo_id}</strong>
        {' '}— {submitted.crew_count} crew member(s) logged
      </p>
      {submitted.photos > 0 && (
        <p className="text-xs text-green-600 font-semibold mb-1">
          ✓ {submitted.photos} photo{submitted.photos!==1?'s':''} uploaded to Drive
        </p>
      )}
      <div className="mb-8 mt-2"><StatusBadge status={submitted.status} /></div>
      <div className="flex gap-3 justify-center flex-wrap">
        <button onClick={reset} className="btn-primary px-6">Submit Another</button>
        <a href="/" className="btn-outline px-6">Dashboard</a>
      </div>
    </div>
  )

  if (loading) return (
    <div className="flex items-center justify-center min-h-[50vh]">
      <div className="flex flex-col items-center gap-4">
        <div className="w-9 h-9 border-[3px] border-slate-200 border-t-navy rounded-full animate-spin" />
        <p className="text-slate-400 text-sm">Loading work orders…</p>
      </div>
    </div>
  )

  if (apiError) return (
    <div className="flex items-center justify-center min-h-[50vh]">
      <div className="text-center max-w-sm">
        <div className="text-4xl mb-3">⚠️</div>
        <p className="text-slate-500 text-sm mb-4">{apiError}</p>
        <button onClick={()=>window.location.reload()} className="btn-outline text-sm px-4 py-2">Reload</button>
      </div>
    </div>
  )

  return (
    <div className="max-w-2xl mx-auto px-4 py-6 space-y-4">

      {/* No-photo confirmation modal */}
      {showPhotoModal && (
        <ConfirmModal
          title="No Photos Selected"
          message="You haven't added any site photos. Photos are important for documentation and verification. Are you sure you want to submit without them?"
          confirmLabel="Submit Without Photos"
          cancelLabel="Go Back & Add Photos"
          danger
          onConfirm={()=>{ setShowPhotoModal(false); doSubmit() }}
          onCancel={()=>setShowPhotoModal(false)}
        />
      )}

      <div>
        <h1 className="text-2xl font-black text-navy">Field Report</h1>
        <p className="text-slate-500 text-sm mt-0.5">Log crew sign-in, work progress, and photos for a Work Order.</p>
      </div>

      {formError && (
        <div className="bg-red-50 border border-red-200 text-red-700 rounded-xl p-4 text-sm">{formError}</div>
      )}

      <form onSubmit={handleSubmit} className="space-y-4">

        {/* 1 · Work Order */}
        <div className="card p-4 space-y-3">
          <p className="section-label">Work Order</p>
          <Field label="Select Work Order" required>
            <select value={selectedWOId} onChange={e=>setSelectedWOId(e.target.value)} className="field-input">
              <option value="">— Choose a Work Order —</option>
              {wos.map(wo=>(
                <option key={wo.id} value={wo.id}>{wo.id} — {wo.location} ({wo.borough})</option>
              ))}
            </select>
          </Field>
          {selectedWO && <WOPanel wo={selectedWO} />}
        </div>

        {/* 2 · Work Details */}
        <div className="card p-4 space-y-4">
          <p className="section-label">Work Details</p>
          <Field label="Date of Work" required>
            <input type="date" value={workDate} onChange={e=>setWorkDate(e.target.value)} className="field-input" />
          </Field>
          <div className="space-y-2">
            <label className="field-label">Marking Types Completed</label>
            <div className="space-y-1.5">
              {markingRows.map((row,i)=>(
                <MarkingRow key={i} idx={i} data={row} onChange={updateMarking} onRemove={removeMarking} />
              ))}
            </div>
            <button type="button" onClick={addMarking} className="btn-ghost text-xs">+ Add Marking Type</button>
            {markingTypesStr && (
              <p className="text-[11px] text-slate-400 bg-slate-50 rounded-lg px-2 py-1.5 font-mono">{markingTypesStr}</p>
            )}
          </div>
          <div className="grid grid-cols-2 gap-3">
            <Field label="SQFT Completed">
              <input type="number" value={sqft} onChange={e=>setSqft(e.target.value)}
                placeholder="0" min="0" inputMode="numeric" className="field-input" />
            </Field>
            <Field label="Paint / Material">
              <input type="text" value={paintMaterial} onChange={e=>setPaintMaterial(e.target.value)}
                placeholder="e.g. White Thermo" className="field-input" />
            </Field>
          </div>
          <Field label="Issues / Notes">
            <textarea value={issues} onChange={e=>setIssues(e.target.value)}
              placeholder="Problems, delays, or anything admin should know…"
              rows={3} className="field-input resize-none" />
          </Field>
        </div>

        {/* 3 · Photos */}
        <div className="card p-4 space-y-3">
          <p className="section-label">Site Photos</p>
          {selectedWOId
            ? <PhotoPicker onChange={setPhotoFiles} />
            : <p className="text-sm text-slate-400 italic">Select a Work Order above to add photos.</p>
          }
        </div>

        {/* 4 · Crew & Signatures */}
        <div className="card p-4 space-y-3">
          <p className="section-label">Crew & Signatures ({crewMembers.length})</p>
          <p className="text-[11px] text-slate-400">
            Each employee signs once at time-in and once at time-out. Signatures will be added to the Sign-In Log.
          </p>
          <div className="space-y-4">
            {crewMembers.map((m,i)=>(
              <CrewCard key={i} idx={i} data={m} onChange={updateCrew} onRemove={removeCrew} />
            ))}
          </div>
          <button type="button" onClick={addCrew} className="btn-ghost text-xs w-full">+ Add Crew Member</button>
        </div>

        {/* 5 · Completion */}
        <div className="card p-4 space-y-3">
          <p className="section-label">Completion</p>
          <Field label="Is this Work Order complete?">
            <YNToggle value={woComplete} onChange={setWoComplete}
              noLabel="No — more work needed" yesLabel="Yes — WO Done ✓" />
          </Field>
        </div>

        {/* Submit */}
        <button type="submit" disabled={submitting} className="btn-primary w-full text-base">
          {submitting ? (submitStep || 'Submitting…') : 'Submit Field Report'}
        </button>

        <div className="h-8" />
      </form>
    </div>
  )
}
