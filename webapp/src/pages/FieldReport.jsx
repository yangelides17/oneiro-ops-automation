import { useState, useEffect, useRef } from 'react'
import StatusBadge  from '../components/StatusBadge'
import SignaturePad from '../components/SignaturePad'
import ConfirmModal from '../components/ConfirmModal'

// ── Constants ─────────────────────────────────────────────────
// Must match the dropdown validation on the Daily Sign-In Data sheet's
// Classification column — LP (Line Person / Crew Chief) and SAT (Stripe
// Assistant Tech). Adding values here without updating the sheet validation
// will cause sign-in submissions to fail with "Invalid Entry".
const CLASSIFICATIONS = ['LP', 'SAT']

// Canonical marking categories for the "Add marking manually" dropdown.
// Mirrors the list in Apps Script setupMarkingItems() / Marking Items col F.
const MARKING_CATEGORIES = [
  // WO Top Table
  'Double Yellow Line', 'Lane Lines', 'Gores', 'Messages', 'Arrows',
  'Solid Lines', 'Rail Road X/Diamond', 'Others',
  // Intersection Grid
  'HVX Crosswalk', 'Stop Msg', 'Stop Line',
  // Page 2 detailed lines
  '4" Line', '6" Line', '8" Line', '12" Line', '16" Line', '24" Line',
  // Page 2 messages
  'Only Msg', 'Bus Msg', 'Bump Msg', 'Custom Msg', '20 MPH Msg',
  // Page 2 railroad
  'Railroad (RR)', 'Railroad (X)',
  // Page 2 arrows
  'L/R Arrow', 'Straight Arrow', 'Combination Arrow',
  // Page 2 misc
  'Speed Hump Markings', 'Shark Teeth 12x18', 'Shark Teeth 24x36',
  // Page 2 bike lane
  'Bike Lane Arrow', 'Bike Lane Symbol', 'Bike Lane Green Bar',
  // MMA
  'Bike Lane', 'Pedestrian Space', 'Bus Lane', 'Ped Stop',
]

const SECTION_HEADERS = {
  'Top Table':          'WO Top Table',
  'Intersection Grid':  'Intersection Grid',
  'Manual':             'Manually Added',
}

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
const newCrew = () => ({ name:'', classification:CLASSIFICATIONS[0], timeIn:'', timeOut:'', hours:'', overtime:'', signatureIn:null, signatureOut:null })
const newManualMarking = () => ({
  category: MARKING_CATEGORIES[0], description: '', intersection: '', direction: '',
  unit: 'SF', quantity: '', color_material: ''
})

// Human label for a marking item — e.g. "Double Yellow Line  •  RECAP FROM HAMILTON PL TO 2ND AV"
// or "2 AV — East HVX" for grid entries.
function itemLabel(item) {
  if (item.section === 'Intersection Grid' && item.intersection) {
    const dirFull = { N: 'North', E: 'East', S: 'South', W: 'West' }[item.direction] || item.direction
    return `${item.intersection} — ${dirFull} ${item.category}`
  }
  return item.description
    ? `${item.category} · ${item.description}`
    : item.category
}

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

// ── Planned marking row (pre-populated from WO scan) ──────────
// Read-only label + editable qty/material. Thermo items hide material.
function MarkingItemRow({ item, onChange }) {
  const isThermo = String(item.work_type || '').toLowerCase() === 'thermo'
  return (
    <div className="border border-slate-200 rounded-lg p-2.5 bg-white space-y-2">
      <div className="flex items-baseline gap-2">
        <span className="flex-1 text-sm font-semibold text-navy leading-tight break-words">
          {itemLabel(item)}
        </span>
        {item.planned && (
          <span className="text-[10px] font-mono uppercase tracking-wide text-slate-400 shrink-0">
            planned: {item.planned}
          </span>
        )}
      </div>
      <div className={`grid gap-2 items-center ${isThermo ? 'grid-cols-[60px_1fr]' : 'grid-cols-[60px_60px_1fr]'}`}>
        <input
          type="number" min="0" inputMode="numeric" placeholder={item.unit || 'SF'}
          value={item.quantity ?? ''}
          onChange={e => onChange(item.item_id, { quantity: e.target.value, _dirty: true })}
          className="field-input text-sm py-2 text-center" />
        <select
          value={item.unit || 'SF'}
          onChange={e => onChange(item.item_id, { unit: e.target.value, _dirty: true })}
          className="field-input text-sm py-2">
          {['SF','LF','EA'].map(u => <option key={u}>{u}</option>)}
        </select>
        {!isThermo && (
          <input
            type="text" placeholder="Color / Material (e.g. White Thermo)"
            value={item.color_material ?? ''}
            onChange={e => onChange(item.item_id, { color_material: e.target.value, _dirty: true })}
            className="field-input text-sm py-2" />
        )}
      </div>
    </div>
  )
}

// ── Manual marking row (not from WO scan) ──────────────────────
function ManualMarkingRow({ idx, data, onChange, onRemove }) {
  return (
    <div className="border border-amber-200 rounded-lg p-2.5 bg-amber-50/50 space-y-2">
      <div className="flex items-center justify-between">
        <span className="text-[10px] font-bold uppercase tracking-wider text-amber-700">Manual entry</span>
        <button type="button" onClick={() => onRemove(idx)}
          className="text-red-400 hover:text-red-600 text-xl leading-none h-6 w-6 flex items-center justify-center">×</button>
      </div>
      <div className="grid grid-cols-2 gap-2">
        <select value={data.category} onChange={e => onChange(idx, { category: e.target.value })}
          className="field-input text-sm py-2">
          {MARKING_CATEGORIES.map(c => <option key={c}>{c}</option>)}
        </select>
        <input type="text" placeholder="Description (optional)"
          value={data.description} onChange={e => onChange(idx, { description: e.target.value })}
          className="field-input text-sm py-2" />
      </div>
      <div className="grid grid-cols-2 gap-2">
        <input type="text" placeholder="Intersection (optional)"
          value={data.intersection} onChange={e => onChange(idx, { intersection: e.target.value })}
          className="field-input text-sm py-2" />
        <select value={data.direction} onChange={e => onChange(idx, { direction: e.target.value })}
          className="field-input text-sm py-2">
          <option value="">No direction</option>
          {['N','E','S','W'].map(d => <option key={d}>{d}</option>)}
        </select>
      </div>
      <div className="grid grid-cols-[60px_60px_1fr] gap-2">
        <input type="number" min="0" inputMode="numeric" placeholder="Qty"
          value={data.quantity} onChange={e => onChange(idx, { quantity: e.target.value })}
          className="field-input text-sm py-2 text-center" />
        <select value={data.unit} onChange={e => onChange(idx, { unit: e.target.value })}
          className="field-input text-sm py-2">
          {['SF','LF','EA'].map(u => <option key={u}>{u}</option>)}
        </select>
        <input type="text" placeholder="Color / Material (leave blank for Thermo)"
          value={data.color_material} onChange={e => onChange(idx, { color_material: e.target.value })}
          className="field-input text-sm py-2" />
      </div>
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
  const [markingItems,  setMarkingItems]   = useState([])   // loaded from /api/wo-markings/:woId
  const [newMarkings,   setNewMarkings]    = useState([])   // manually added this session
  const [markingsLoading, setMarkingsLoading] = useState(false)
  const [issues,        setIssues]         = useState('')
  const [crewMembers,   setCrewMembers]    = useState([newCrew()])
  const [woComplete,    setWoComplete]     = useState('no')
  const [photoFiles,    setPhotoFiles]     = useState([])   // File objects, uploaded on submit
  const [crewLeaderName,      setCrewLeaderName]      = useState('')
  const [crewLeaderSignature, setCrewLeaderSignature] = useState(null)

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

  // Load pre-populated marking items from the WO scan whenever the
  // selected WO changes. Reset any in-progress manual additions too.
  useEffect(() => {
    if (!selectedWOId) {
      setMarkingItems([])
      setNewMarkings([])
      return
    }
    setMarkingsLoading(true)
    fetch(`/api/wo-markings/${encodeURIComponent(selectedWOId)}`)
      .then(r => r.json())
      .then(d => {
        if (d.error) throw new Error(d.error)
        // Sort by sort_order so the list reads top-to-bottom against the paper WO.
        const items = (d.items || []).slice().sort((a, b) => (a.sort_order ?? 0) - (b.sort_order ?? 0))
        setMarkingItems(items)
        setNewMarkings([])
      })
      .catch(e => {
        console.error('Failed to load marking items:', e)
        setMarkingItems([])
      })
      .finally(() => setMarkingsLoading(false))
  }, [selectedWOId])

  // Planned marking updates — in-place edits, with a _dirty flag so we
  // only send touched rows back.
  const updateMarkingItem = (itemId, patch) =>
    setMarkingItems(list => list.map(it => it.item_id === itemId ? { ...it, ...patch } : it))

  // Manual markings
  const addManualMarking = () => setNewMarkings(list => [...list, newManualMarking()])
  const updateManualMarking = (idx, patch) =>
    setNewMarkings(list => list.map((x, j) => j === idx ? { ...x, ...patch } : x))
  const removeManualMarking = (idx) =>
    setNewMarkings(list => list.filter((_, j) => j !== idx))

  // Crew rows
  const updateCrew = (i, updated) => setCrewMembers(l=>l.map((m,j)=>j===i?updated:m))
  const addCrew    = () => setCrewMembers(l=>[...l,newCrew()])
  const removeCrew = (i) => setCrewMembers(l=>l.filter((_,j)=>j!==i))

  // Group planned items by section for rendering (Top Table → Grid → other)
  const itemsBySection = markingItems.reduce((acc, it) => {
    const sec = it.section || 'Other'
    ;(acc[sec] = acc[sec] || []).push(it)
    return acc
  }, {})
  const SECTION_ORDER = ['Top Table', 'Intersection Grid', 'Manual', 'Other']

  // Infer WO work type from any loaded planned item (Thermo beats MMA if
  // mixed). Falls back to blank if no items are loaded yet.
  const inferredWorkType = markingItems.some(i => String(i.work_type).toLowerCase() === 'thermo')
    ? 'Thermo'
    : markingItems.some(i => String(i.work_type).toLowerCase() === 'mma') ? 'MMA' : ''

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
      // Only send touched planned items. "Dirty" = user edited qty/unit/material.
      const markingUpdates = markingItems
        .filter(i => i._dirty)
        .map(i => ({
          item_id:        i.item_id,
          quantity:       i.quantity === '' || i.quantity == null ? null : parseFloat(i.quantity),
          color_material: (i.color_material || '').trim(),
          unit:           i.unit || 'SF',
        }))

      // All manual adds that have at least a category are included.
      const markingNew = newMarkings
        .filter(n => (n.category || '').trim())
        .map(n => ({
          category:       (n.category || '').trim(),
          description:    (n.description || '').trim(),
          intersection:   (n.intersection || '').trim(),
          direction:      (n.direction || '').trim(),
          unit:           n.unit || 'SF',
          quantity:       n.quantity === '' || n.quantity == null ? null : parseFloat(n.quantity),
          color_material: (n.color_material || '').trim(),
        }))

      const res = await fetch('/api/field-report', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          wo_id:       selectedWOId,
          date:        workDate,
          wo_complete: woComplete==='yes',
          work_type:   inferredWorkType,
          marking_updates: markingUpdates,
          marking_new:     markingNew,
          issues:          issues.trim(),
          photos_uploaded: photosUploaded,
          crew: validCrew.map(m=>({
            name:           m.name.trim(),
            classification: m.classification,
            time_in:        fmt24to12(m.timeIn),
            time_out:       fmt24to12(m.timeOut),
            hours:          parseFloat(m.hours)||0,
            overtime:       parseFloat(m.overtime)||0,
            // Base64 PNG data URLs from the SignaturePad canvas. Apps Script
            // forwards these verbatim into the Sign-In Logs JSON — the Python
            // worker decodes + embeds them into the filled PDF. Never archived.
            sig_in_b64:     m.signatureIn  || '',
            sig_out_b64:    m.signatureOut || ''
          })),
          // Crew-leader block (signs at the bottom of the Sign-In Log)
          contractor_name:           crewLeaderName.trim(),
          contractor_title:          'Crew Leader',
          contractor_signature_b64:  crewLeaderSignature || ''
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
    setMarkingItems([]); setNewMarkings([]); setIssues('')
    setCrewMembers([newCrew()]); setWoComplete('no'); setPhotoFiles([]); setFormError('')
    setCrewLeaderName(''); setCrewLeaderSignature(null)
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

          {/* ── Marking Items ─────────────────────────────────── */}
          <div className="space-y-2">
            <div className="flex items-baseline justify-between">
              <label className="field-label">Marking Items</label>
              {inferredWorkType && (
                <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400">
                  {inferredWorkType}
                </span>
              )}
            </div>

            {!selectedWOId && (
              <p className="text-xs text-slate-400 italic">
                Select a Work Order above to see its marking items.
              </p>
            )}

            {selectedWOId && markingsLoading && (
              <p className="text-xs text-slate-400 italic">Loading marking items…</p>
            )}

            {selectedWOId && !markingsLoading && markingItems.length === 0 && newMarkings.length === 0 && (
              <p className="text-xs text-slate-400 italic bg-slate-50 rounded-lg px-3 py-2">
                No items pre-populated from the scan. Use "Add marking manually" below to log what was done.
              </p>
            )}

            {/* Planned items, grouped by section */}
            {SECTION_ORDER.map(sec => {
              const rows = itemsBySection[sec] || []
              if (!rows.length) return null
              return (
                <div key={sec} className="space-y-1.5">
                  <p className="text-[10px] font-bold uppercase tracking-wider text-slate-400 mt-2">
                    {SECTION_HEADERS[sec] || sec}
                  </p>
                  {rows.map(item => (
                    <MarkingItemRow key={item.item_id} item={item} onChange={updateMarkingItem} />
                  ))}
                </div>
              )
            })}

            {/* Manual additions */}
            {newMarkings.length > 0 && (
              <div className="space-y-1.5">
                <p className="text-[10px] font-bold uppercase tracking-wider text-amber-700 mt-2">
                  Manually Added
                </p>
                {newMarkings.map((nm, i) => (
                  <ManualMarkingRow key={i} idx={i} data={nm}
                    onChange={updateManualMarking} onRemove={removeManualMarking} />
                ))}
              </div>
            )}

            {selectedWOId && (
              <button type="button" onClick={addManualMarking} className="btn-ghost text-xs mt-1">
                + Add marking manually
              </button>
            )}
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

        {/* 5 · Crew Leader Sign-Off */}
        <div className="card p-4 space-y-3">
          <p className="section-label">Crew Leader Sign-Off</p>
          <p className="text-[11px] text-slate-400">
            The crew leader signs at the bottom of the Sign-In Log confirming the day's crew and hours.
          </p>
          <Field label="Crew Leader Name">
            <input type="text" value={crewLeaderName} autoCapitalize="words" placeholder="First Last"
              onChange={e=>setCrewLeaderName(e.target.value)} className="field-input" />
          </Field>
          <SignaturePad
            label="Crew Leader Signature"
            onChange={setCrewLeaderSignature}
          />
        </div>

        {/* 6 · Completion */}
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
