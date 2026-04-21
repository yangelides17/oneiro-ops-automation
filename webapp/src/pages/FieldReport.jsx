import { useState, useEffect, useRef } from 'react'
import StatusBadge      from '../components/StatusBadge'
import SignaturePad     from '../components/SignaturePad'
import ConfirmModal     from '../components/ConfirmModal'
import MarkingFormModal from '../components/MarkingFormModal'
import RowKebab         from '../components/RowKebab'
import {
  UNIT_OPTIONS, unitForCategory, unitIsLocked,
  pickLayout, rowIsCompletable,
} from '../lib/markingCategories'

// ── Constants ─────────────────────────────────────────────────
// Must match the dropdown validation on the Daily Sign-In Data sheet's
// Classification column — LP (Line Person / Crew Chief) and SAT (Stripe
// Assistant Tech). Adding values here without updating the sheet validation
// will cause sign-in submissions to fail with "Invalid Entry".
const CLASSIFICATIONS = ['LP', 'SAT']

const SECTION_HEADERS = {
  'Top Table':          'WO Top Table',
  'Intersection Grid':  'Intersection Grid',
  'Manual':             'Manually Added',
}

// pickLayout / rowIsCompletable / unit helpers live in
// ../lib/markingCategories — single source of truth shared with the
// Add/Edit modal.

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
// Day-of-week-aware OT calc. Rule:
//   Sat/Sun → every hour is overtime.
//   Mon–Fri → hours over 8 are overtime.
// workDateIso is the Date of Work as "YYYY-MM-DD" (from <input type="date">).
// We construct a local-parts Date so there's no UTC-shift in the DOW.
const calcHours = (tin, tout, workDateIso) => {
  if (!tin || !tout) return { hours: '', overtime: '' }
  const [ih,im] = tin.split(':').map(Number)
  const [oh,om] = tout.split(':').map(Number)
  const mins = (oh*60+om) - (ih*60+im)
  if (mins <= 0) return { hours: '', overtime: '' }
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
const newCrew = () => ({ name:'', classification:CLASSIFICATIONS[0], timeIn:'', timeOut:'', hours:'', overtime:'', signatureIn:null, signatureOut:null })

// ── HEIC → JPEG converter ─────────────────────────────────────
// iPhones ship HEIC by default. Browsers can't decode HEIC for
// <img> preview or canvas, so we convert to JPEG the moment a file
// is picked. The converted File is what flows through the rest of
// the pipeline (preview + compressImage + upload), so no other code
// needs to know HEIC exists. heic2any is dynamically imported so
// the 180 KB WASM payload only loads when a HEIC actually appears.
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
    const converted = await heic2any({ blob: file, toType: 'image/jpeg', quality: 0.9 })
    const blob      = Array.isArray(converted) ? converted[0] : converted
    const baseName  = (file.name || 'photo').replace(/\.(heic|heif)$/i, '')
    return new File([blob], `${baseName}.jpg`, { type: 'image/jpeg' })
  } catch (err) {
    // Upstream callers can still upload the original — Drive will store
    // the HEIC bytes. Only the preview stays broken.
    console.warn('HEIC → JPEG conversion failed; keeping original:', err)
    return file
  }
}

// ── Image compression helper ──────────────────────────────────
// Phone JPEGs land at 3-10 MB straight from the camera. Shipping those
// raw through the Apps Script proxy is the single biggest source of
// submit-time latency. We resize to 2048px long-edge and re-encode at
// JPEG q=0.85 before upload. Typical output is 300-600 KB with no
// visible quality loss. Non-JPEG/PNG inputs (or already-small files)
// pass through untouched.
const COMPRESS_MAX_EDGE = 2048
const COMPRESS_QUALITY  = 0.85
const COMPRESS_SKIP_BELOW_BYTES = 500 * 1024   // <500 KB — not worth recompressing
async function compressImage(file) {
  const type = String(file?.type || '').toLowerCase()
  if (!file) return file
  // Only attempt to compress JPEG/PNG (HEIC etc. need a library we don't ship)
  if (type !== 'image/jpeg' && type !== 'image/png') return file
  if (file.size <= COMPRESS_SKIP_BELOW_BYTES) return file

  let bitmap
  try {
    bitmap = await createImageBitmap(file)
  } catch {
    return file  // can't decode — fall back to original
  }

  const { width, height } = bitmap
  const longEdge = Math.max(width, height)
  const scale = longEdge > COMPRESS_MAX_EDGE ? COMPRESS_MAX_EDGE / longEdge : 1
  const w = Math.round(width * scale)
  const h = Math.round(height * scale)

  const canvas = typeof OffscreenCanvas !== 'undefined'
    ? new OffscreenCanvas(w, h)
    : Object.assign(document.createElement('canvas'), { width: w, height: h })
  const ctx = canvas.getContext('2d')
  if (!ctx) return file
  ctx.drawImage(bitmap, 0, 0, w, h)
  bitmap.close?.()

  // Prefer canvas.convertToBlob (OffscreenCanvas) for async; fall back to toBlob
  const blob = canvas.convertToBlob
    ? await canvas.convertToBlob({ type: 'image/jpeg', quality: COMPRESS_QUALITY })
    : await new Promise(r => canvas.toBlob(r, 'image/jpeg', COMPRESS_QUALITY))
  if (!blob || blob.size >= file.size) return file   // compression made it larger somehow

  // Give the blob a filename so the server sees a sensible name
  const baseName = (file.name || 'photo').replace(/\.(png|jpg|jpeg|heic|heif|webp)$/i, '')
  return new File([blob], `${baseName}.jpg`, { type: 'image/jpeg' })
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

// ── Planned marking row ───────────────────────────────────────
// Single-line layout per section/category so rows align vertically like a
// grid. Layout is driven by `pickLayout(item)`:
//   grid    → [Type][Intersection][Direction][Qty][Unit]
//   mma     → [Type][Color/Material][Qty][Unit]
//   default → [Type][Qty][Unit]
//
// Inline fields save on commit (blur / Enter / dropdown change) — not on
// every keystroke — to avoid persisting intermediate values while the
// user is still typing. Once Status='Completed', all inline fields are
// read-only; use the kebab → Edit to reopen editing via the modal.
//
// Leftmost cell: checkbox (only rendered in bulk-select mode).
// Rightmost cell: kebab or per-row spinner.
function MarkingItemRow({
  item, selected, saving, bulkMode,
  onToggleSelect, onLocalChange, onCommit,
  onEdit, onDelete, onStartBulk,
}) {
  const layout = pickLayout(item)
  const locked = item.status === 'Completed'

  const INPUT = 'field-input text-sm py-2'
  const RO    = INPUT + ' bg-slate-50 text-slate-500 cursor-default focus:ring-0'
  const LOCK  = INPUT + ' bg-slate-50 text-slate-500 cursor-not-allowed focus:ring-0'
  const QTY_LIVE   = INPUT + ' text-center'
  const QTY_LOCKED = LOCK  + ' text-center'

  // onChange handlers. Text inputs update local state only; commits
  // happen in onBlur/onKeyDown. Unit (dropdown) commits immediately.
  const onLocalText  = (field) => (e) => onLocalChange(item.item_id, field, e.target.value)
  const onCommitText = (field) => (e) => onCommit(item.item_id, field, e.target.value)
  const onDropdown   = (field) => (e) => {
    onLocalChange(item.item_id, field, e.target.value)
    onCommit(item.item_id, field, e.target.value)
  }
  const onEnterBlur = (e) => { if (e.key === 'Enter') e.target.blur() }

  const CategoryBox = (
    <input type="text" readOnly value={item.category ?? ''}
      placeholder="Marking Type" className={RO} />
  )
  const QtyBox = (
    <input
      type="number" min="0" inputMode="numeric" placeholder="Qty"
      value={item.quantity == null ? '' : item.quantity}
      onChange={onLocalText('quantity')}
      onBlur={onCommitText('quantity')}
      onKeyDown={onEnterBlur}
      readOnly={locked}
      className={locked ? QTY_LOCKED : QTY_LIVE}
    />
  )
  // Unit is derived from Marking Type for every category except "Others",
  // which stays user-pickable. For locked categories, show the unit as
  // read-only text (no dropdown chrome).
  const unitLocked = unitIsLocked(item.category)
  const derivedUnit = unitForCategory(item.category) || item.unit || ''
  const UnitBox = unitLocked
    ? <div className="text-sm py-2 text-center font-semibold text-slate-500">
        {derivedUnit}
      </div>
    : <select value={item.unit || 'EA'}
        onChange={onDropdown('unit')}
        disabled={locked}
        className={locked ? LOCK : INPUT}>
        {UNIT_OPTIONS.map(u => <option key={u}>{u}</option>)}
      </select>
  const ColorBox = (
    <input type="text" placeholder="Color / Material"
      value={item.color_material ?? ''}
      onChange={onLocalText('color_material')}
      onBlur={onCommitText('color_material')}
      onKeyDown={onEnterBlur}
      readOnly={locked}
      className={locked ? LOCK : INPUT}
    />
  )
  const CheckboxCell = bulkMode ? (
    <input type="checkbox" checked={selected}
      onChange={e => onToggleSelect(item.item_id, e.target.checked)}
      className="w-4 h-4 accent-navy cursor-pointer" />
  ) : null
  const ActionCell = saving
    ? <Spinner />
    : <RowKebab items={[
        { label: 'Edit',        onClick: () => onEdit(item) },
        { label: 'Delete',      onClick: () => onDelete(item), danger: true },
        { label: 'Bulk delete', onClick: () => onStartBulk() },
      ]} />

  // Grid templates — [checkbox?] + fields + [action]. Checkbox col is
  // 0px when bulk mode is off (hidden entirely).
  const CB  = bulkMode ? '24px ' : ''
  const END = ' 28px'
  const tpl = layout === 'grid'
    ? `${CB}1fr 110px 64px 90px 56px${END}`
    : layout === 'mma'
      ? `${CB}1fr 1fr 90px 56px${END}`
      : `${CB}1fr 90px 56px${END}`

  const DescNote = item.description
    ? <p className={`text-[11px] text-slate-400 ${bulkMode ? 'pl-9' : 'pl-1'}`}>Note: {item.description}</p>
    : null

  return (
    <div className="space-y-0.5">
      <div className="grid items-center gap-2" style={{ gridTemplateColumns: tpl }}>
        {CheckboxCell}
        {CategoryBox}

        {layout === 'grid' && (
          <>
            <input type="text" readOnly value={item.intersection ?? ''}
              placeholder="Intersection" className={RO} />
            <input type="text" readOnly value={item.direction ?? ''}
              placeholder="Direction" className={RO + ' text-center'} />
          </>
        )}
        {layout === 'mma' && ColorBox}

        {QtyBox}
        {UnitBox}
        {ActionCell}
      </div>
      {/* Top-table / MMA items may carry a description; grid items don't. */}
      {layout !== 'grid' && DescNote}
    </div>
  )
}

// ── Tiny inline spinner ───────────────────────────────────────
function Spinner() {
  return (
    <div className="w-7 h-7 flex items-center justify-center">
      <div className="w-4 h-4 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
    </div>
  )
}

// ── Crew card (with signature pads) ──────────────────────────
function CrewCard({ idx, data, onChange, onRemove, workDate }) {
  const handleTime = (field, val) => {
    const tin  = field==='timeIn'  ? val : data.timeIn
    const tout = field==='timeOut' ? val : data.timeOut
    onChange(idx, { ...data, [field]: val, ...calcHours(tin, tout, workDate) })
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
// Items tracked per-photo:
//   { id, name, status: 'processing'|'ready'|'error', file, previewUrl, error }
// HEIC files get converted to JPEG the moment they're picked so the
// preview works and the submit-time compression can run against a
// format the browser knows how to decode.
function PhotoPicker({ onChange }) {
  const fileRef = useRef(null)
  const [items, setItems] = useState([])

  // Emit the final file list to the parent whenever items change.
  // Errored items are omitted — the user can see them in the UI and
  // decide to retry or remove. Processing items still emit their
  // original file so the parent "has" something; the submit flow
  // blocks on processing state via a separate signal below.
  useEffect(() => {
    onChange(items
      .filter(i => i.status !== 'error')
      .map(i => i.file))
  }, [items])   // eslint-disable-line react-hooks/exhaustive-deps

  // Clean up object URLs on unmount so we don't leak blobs
  useEffect(() => () => {
    items.forEach(i => { if (i.previewUrl) URL.revokeObjectURL(i.previewUrl) })
  }, [])   // eslint-disable-line react-hooks/exhaustive-deps

  const addFiles = (e) => {
    const added = Array.from(e.target.files || [])
    e.target.value = ''
    if (added.length === 0) return

    const newItems = added.map(f => ({
      id:         `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      name:       f.name || 'photo',
      status:     isHeic(f) ? 'processing' : 'ready',
      file:       f,
      previewUrl: isHeic(f) ? null : URL.createObjectURL(f),
      error:      null,
    }))
    setItems(prev => [...prev, ...newItems])

    // Convert HEICs in the background — non-HEICs are already 'ready'
    newItems.forEach(item => {
      if (item.status === 'ready') return
      convertHeicToJpeg(item.file).then(processed => {
        const converted   = processed !== item.file
        const previewUrl  = URL.createObjectURL(processed)
        setItems(prev => prev.map(p => p.id === item.id
          ? { ...p, file: processed, previewUrl, status: converted ? 'ready' : 'error',
              error: converted ? null : 'Couldn\u2019t convert HEIC — preview unavailable' }
          : p))
      }).catch(err => {
        setItems(prev => prev.map(p => p.id === item.id
          ? { ...p, status: 'error', error: err?.message || 'Failed to prepare photo' }
          : p))
      })
    })
  }

  const remove = (id) => setItems(prev => {
    const target = prev.find(p => p.id === id)
    if (target?.previewUrl) URL.revokeObjectURL(target.previewUrl)
    return prev.filter(p => p.id !== id)
  })

  const processingCount = items.filter(i => i.status === 'processing').length
  const readyCount      = items.filter(i => i.status === 'ready').length
  const errorCount      = items.filter(i => i.status === 'error').length

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
      {items.length > 0 && (
        <>
          <div className="flex flex-wrap gap-2">
            {items.map(item => (
              <div key={item.id} className="relative group">
                {item.status === 'ready' && item.previewUrl && (
                  <img src={item.previewUrl} alt={item.name}
                    className="w-16 h-16 object-cover rounded-lg border border-slate-200" />
                )}
                {item.status === 'processing' && (
                  <div className="w-16 h-16 rounded-lg border border-slate-200 bg-slate-50
                                  flex flex-col items-center justify-center gap-0.5">
                    <div className="w-4 h-4 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
                    <span className="text-[8px] text-slate-400 uppercase tracking-wider">HEIC</span>
                  </div>
                )}
                {item.status === 'error' && (
                  <div className="w-16 h-16 rounded-lg border border-red-200 bg-red-50
                                  flex items-center justify-center text-red-500 text-lg font-bold"
                       title={item.error || 'Error'}>!</div>
                )}
                <button type="button" onClick={()=>remove(item.id)}
                  className="absolute -top-1.5 -right-1.5 w-5 h-5 bg-red-500 text-white rounded-full
                             text-xs font-bold flex items-center justify-center
                             opacity-0 group-hover:opacity-100 transition-opacity">×</button>
              </div>
            ))}
          </div>
          <p className="text-[11px] font-semibold flex flex-wrap gap-x-3 gap-y-1">
            {readyCount > 0 && (
              <span className="text-green-700">
                ✓ {readyCount} ready
              </span>
            )}
            {processingCount > 0 && (
              <span className="text-slate-500">
                Preparing {processingCount}…
              </span>
            )}
            {errorCount > 0 && (
              <span className="text-red-600">
                ✕ {errorCount} couldn&apos;t prepare
              </span>
            )}
          </p>
        </>
      )}

      {items.length === 0 && (
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
  const [markingsLoading, setMarkingsLoading] = useState(false)
  const [issues,        setIssues]         = useState('')
  const [crewMembers,   setCrewMembers]    = useState([newCrew()])
  const [woComplete,    setWoComplete]     = useState('no')
  const [photoFiles,    setPhotoFiles]     = useState([])   // File objects, uploaded on submit
  const [crewLeaderName,      setCrewLeaderName]      = useState('')
  const [crewLeaderSignature, setCrewLeaderSignature] = useState(null)

  // Per-row UI state for the Marking Items list
  const [bulkMode,      setBulkMode]      = useState(false)
  const [selectedIds,   setSelectedIds]   = useState(new Set())
  const [rowSaving,     setRowSaving]     = useState(new Set())
  const [formModal,     setFormModal]     = useState(null)  // {mode, item} or null
  const [deleteConfirm, setDeleteConfirm] = useState(null)  // {ids:[...], label?} or null
  const [rowError,      setRowError]      = useState('')
  // In-flight saves mirror of rowSaving for awaitable access from doSubmit.
  const inFlightRef = useRef(new Set())
  // Ref on the error banner so we can scroll it into view when it appears —
  // the Submit button is at the bottom of the form, so a new error would
  // otherwise be off-screen at the top.
  const errorRef = useRef(null)

  // UI state
  const [formError,      setFormError]      = useState('')
  // Bumps every time we set an error — even if the message is the same
  // string the user already saw — so the scroll-to-error effect fires on
  // repeat submits too.
  const [errorNonce,     setErrorNonce]     = useState(0)
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
  // selected WO changes.
  useEffect(() => {
    if (!selectedWOId) {
      setMarkingItems([])
      setSelectedIds(new Set())
      return
    }
    setMarkingsLoading(true)
    fetch(`/api/wo-markings/${encodeURIComponent(selectedWOId)}`)
      .then(r => r.json())
      .then(d => {
        if (d.error) throw new Error(d.error)
        // Preserve sheet row order (= insertion order from the scan + any
        // earlier manual adds) — no client-side sorting needed.
        setMarkingItems(d.items || [])
        setSelectedIds(new Set())
      })
      .catch(e => {
        console.error('Failed to load marking items:', e)
        setMarkingItems([])
      })
      .finally(() => setMarkingsLoading(false))
  }, [selectedWOId])

  // Scroll the form-level error banner into view every time it's
  // raised — including a repeat submit that produces the same error
  // string (keyed on errorNonce, not formError, so React fires the
  // effect even if the message is identical to the last one).
  useEffect(() => {
    if (formError && errorRef.current) {
      errorRef.current.scrollIntoView({ behavior: 'smooth', block: 'start' })
    }
  }, [errorNonce])

  // Set a new form-level error AND bump the nonce so the scroll effect
  // always fires, even if `msg` matches the previously-shown error.
  const raiseError = (msg) => {
    setFormError(msg)
    setErrorNonce(n => n + 1)
  }

  // ── Per-row save helpers ──────────────────────────────────────
  const markSaving = (itemId, on) => {
    if (on) inFlightRef.current.add(itemId)
    else    inFlightRef.current.delete(itemId)
    setRowSaving(prev => {
      const next = new Set(prev)
      if (on) next.add(itemId); else next.delete(itemId)
      return next
    })
  }

  const refetchMarkings = async () => {
    if (!selectedWOId) return
    try {
      const r = await fetch(`/api/wo-markings/${encodeURIComponent(selectedWOId)}`)
      const d = await r.json()
      if (!d.error) setMarkingItems(d.items || [])
    } catch { /* non-fatal */ }
  }

  // Issue a PATCH for one field. On success, only merge back the fields
  // we explicitly patched plus status/date_completed (server-derived) —
  // so fields the user might still be typing into don't get clobbered.
  const saveField = async (itemId, patch) => {
    markSaving(itemId, true)
    setRowError('')
    try {
      const res = await fetch(`/api/marking-items/${encodeURIComponent(itemId)}`, {
        method:  'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify(patch),
      })
      const data = await res.json()
      if (data.error) throw new Error(data.error)
      if (data.item) {
        setMarkingItems(list => list.map(i => {
          if (i.item_id !== itemId) return i
          const merged = {
            ...i,
            status:         data.item.status,
            date_completed: data.item.date_completed,
          }
          Object.keys(patch).forEach(k => { merged[k] = data.item[k] })
          return merged
        }))
      }
    } catch (err) {
      console.error('save field failed', err)
      setRowError(`Couldn't save: ${err.message}`)
      refetchMarkings()
    } finally {
      markSaving(itemId, false)
    }
  }

  // Optimistic local update — does NOT save. Called as the user types.
  const onRowLocalChange = (itemId, field, value) => {
    setMarkingItems(list => list.map(i =>
      i.item_id === itemId ? { ...i, [field]: value } : i
    ))
  }

  // Commit a value to the server. Called on blur/Enter for text inputs
  // and on-change for dropdowns. Skips the PATCH if the value already
  // matches what's in state (i.e. no real change).
  const onRowCommit = (itemId, field, value) => {
    // Coerce quantity to a number for comparison / send
    const normalized = field === 'quantity'
      ? (value === '' || value == null ? null : parseFloat(value))
      : value
    saveField(itemId, { [field]: normalized })
  }

  // Wait for every in-flight save to finish — called before the
  // Field Report submit so finalize/rollup operate on fresh Drive data.
  const waitForSaves = async () => {
    // Give any onBlur handler currently mid-execution a tick to register.
    await new Promise(r => setTimeout(r, 25))
    while (inFlightRef.current.size > 0) {
      await new Promise(r => setTimeout(r, 50))
    }
  }

  // ── Selection / bulk helpers ──────────────────────────────────
  const toggleSelect = (id, on) => {
    setSelectedIds(prev => {
      const next = new Set(prev)
      if (on) next.add(id); else next.delete(id)
      return next
    })
  }
  const selectAll   = () => setSelectedIds(new Set(markingItems.map(i => i.item_id)))
  const clearSelection = () => setSelectedIds(new Set())
  const enterBulkMode  = () => {
    setBulkMode(true)
    setSelectedIds(new Set())  // no pre-selection per user preference
  }
  const exitBulkMode   = () => {
    setBulkMode(false)
    setSelectedIds(new Set())
  }

  // ── Delete (single or bulk) ───────────────────────────────────
  const doDelete = async (ids) => {
    if (!ids.length) return
    try {
      const res = await fetch('/api/marking-items', {
        method:  'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ item_ids: ids }),
      })
      const data = await res.json()
      if (data.error) throw new Error(data.error)
      const deleted = new Set(data.deleted || ids)
      setMarkingItems(list => list.filter(i => !deleted.has(i.item_id)))
      setSelectedIds(prev => {
        const next = new Set(prev)
        deleted.forEach(id => next.delete(id))
        // If bulk mode was on and all selected items are now gone,
        // collapse back out of bulk mode so the header disappears.
        if (bulkMode && next.size === 0) {
          setBulkMode(false)
        }
        return next
      })
      setRowError('')
    } catch (err) {
      console.error('delete failed', err)
      setRowError(`Delete failed: ${err.message}`)
    } finally {
      setDeleteConfirm(null)
    }
  }

  // Crew rows
  const updateCrew = (i, updated) => setCrewMembers(l=>l.map((m,j)=>j===i?updated:m))
  const addCrew    = () => setCrewMembers(l=>[...l,newCrew()])
  const removeCrew = (i) => setCrewMembers(l=>l.filter((_,j)=>j!==i))

  // When the Date of Work changes, re-run the day-of-week OT rule
  // for every crew member whose times are already filled in. Without
  // this, flipping the date from Fri → Sat (or vice-versa) would leave
  // stale OT splits from before the date change.
  useEffect(() => {
    setCrewMembers(list => list.map(m => {
      if (!m.timeIn || !m.timeOut) return m
      const { hours, overtime } = calcHours(m.timeIn, m.timeOut, workDate)
      if (m.hours === hours && m.overtime === overtime) return m
      return { ...m, hours, overtime }
    }))
  }, [workDate])

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

    // Step 0 — make sure any focused input flushes its save, then wait
    // for all in-flight PATCHes to land so finalize/rollup on the server
    // read the freshest Marking Items state.
    setSubmitStep('Saving pending edits…')
    if (typeof document !== 'undefined' && document.activeElement?.blur) {
      document.activeElement.blur()
    }
    await waitForSaves()

    // Step 1 — compress + upload photos in parallel. Any upload error
    // blocks the submit; we surface it to the user instead of silently
    // proceeding. Compression typically cuts each photo from 3-10 MB
    // down to 300-600 KB.
    // Block if a HEIC is still being converted in PhotoPicker —
    // otherwise we'd ship the raw HEIC up through the proxy.
    if (photoFiles.some(f => isHeic(f))) {
      setSubmitStep('')
      raiseError('HEIC photos are still being prepared. Wait a moment and try again.')
      setSubmitting(false)
      return
    }
    let photosUploaded = false
    if (photoFiles.length > 0) {
      setSubmitStep(`Preparing ${photoFiles.length} photo${photoFiles.length === 1 ? '' : 's'}…`)
      let compressed
      try {
        compressed = await Promise.all(photoFiles.map(compressImage))
      } catch (err) {
        setSubmitStep('')
        raiseError('Couldn\u2019t prepare photos for upload. Try again.')
        setSubmitting(false)
        return
      }

      setSubmitStep(`Uploading ${photoFiles.length} photo${photoFiles.length === 1 ? '' : 's'}…`)
      const results = await Promise.all(compressed.map(async (file) => {
        const form = new FormData()
        form.append('photo', file)
        form.append('wo_id', selectedWOId)
        try {
          const res  = await fetch('/api/upload-photo', { method: 'POST', body: form })
          const data = await res.json().catch(() => ({}))
          if (!res.ok || data.error) return { ok: false, err: data.error || `HTTP ${res.status}` }
          return { ok: true }
        } catch (err) {
          return { ok: false, err: err?.message || 'Network error' }
        }
      }))

      const failures = results.filter(r => !r.ok)
      if (failures.length > 0) {
        setSubmitStep('')
        const detail = failures[0].err ? ` (${failures[0].err})` : ''
        raiseError(`${failures.length} of ${photoFiles.length} photo${photoFiles.length === 1 ? '' : 's'} failed to upload${detail}. Fix and try again.`)
        setSubmitting(false)
        return
      }
      photosUploaded = true
    }

    // Step 2 — submit field report
    setSubmitStep('Submitting report…')
    // Shared payload — used by submit AND by the background finalize
    // call so signatures don't get shipped twice to the browser.
    const reportBody = {
      wo_id:       selectedWOId,
      date:        workDate,
      wo_complete: woComplete==='yes',
      work_type:   inferredWorkType,
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
    }

    try {
      // Marking Items are already live-persisted via per-row CRUD
      // endpoints. Submit only sends WO-level data + crew + signatures.
      const res = await fetch('/api/field-report', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify(reportBody),
      })
      const data = await res.json()
      if (data.error) throw new Error(data.error)
      setSubmitted({ wo_id:data.wo_id, status:data.status, crew_count:validCrew.length, photos:photoFiles.length })

      // Fire-and-forget: kick off Sign-In + CFR JSON generation on the
      // server after the user's already seen the success screen. Any
      // failure lands in the Automation Log sheet — not this UI —
      // because the report data itself is already safely persisted.
      fetch('/api/field-report/finalize', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify(reportBody),
      }).catch(err => {
        // Network-level failure reaching the proxy. Log locally; user
        // doesn't see it because the submit itself succeeded.
        console.warn('field-report finalize request failed:', err)
      })
    } catch (err) {
      raiseError('Submission failed: ' + err.message)
    } finally {
      setSubmitting(false)
      setSubmitStep('')
    }
  }

  // ── Form submit — validate then guard → doSubmit ──────────
  function handleSubmit(e) {
    e.preventDefault()
    setFormError('')
    if (!selectedWOId)  { raiseError('Please select a Work Order.'); return }
    if (!workDate)      { raiseError('Please enter the date of work.'); return }
    const valid = crewMembers.filter(m=>m.name.trim())
    if (!valid.length)  { raiseError('Please add at least one crew member.'); return }
    for (const m of valid) {
      if (!m.timeIn||!m.timeOut) {
        raiseError(`Enter Time In and Time Out for ${m.name||'all crew members'}.`); return
      }
    }
    // If the crew is declaring the WO complete, every Marking Item must
    // be completable (already Completed, or Pending with the required
    // fields filled in). Anything else is either missing a measurement
    // or needs to be deleted — block submit with a specific list.
    if (woComplete === 'yes' && markingItems.length > 0) {
      const bad = markingItems.filter(i => !rowIsCompletable(i))
      if (bad.length > 0) {
        const labels = bad.slice(0, 6).map(i => {
          const base = i.intersection ? `${i.intersection} ${i.direction || ''} ${i.category}`.trim()
                                      : i.category
          return `• ${base}`
        }).join('\n')
        const more = bad.length > 6 ? `\n…and ${bad.length - 6} more` : ''
        raiseError(
          `Can't mark WO complete — ${bad.length} marking item${bad.length===1?'':'s'} ${bad.length===1?'is':'are'} missing required fields ` +
          `(Quantity + Unit${/* MMA color hint */''}, plus Color/Material for MMA items).\n\n${labels}${more}\n\n` +
          `Fill in the missing values or delete the row, then re-submit.`
        )
        return
      }
    }
    // Guard: no photos selected
    if (photoFiles.length===0) { setShowPhotoModal(true); return }
    doSubmit()
  }

  // ── Reset ─────────────────────────────────────────────────
  function reset() {
    setSubmitted(null); setSelectedWOId(''); setWorkDate(isoToday())
    setMarkingItems([]); setSelectedIds(new Set()); setRowSaving(new Set())
    setBulkMode(false); inFlightRef.current.clear()
    setIssues(''); setRowError('')
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

      {/* Marking Items — Add / Edit form modal */}
      {formModal && (
        <MarkingFormModal
          mode={formModal.mode}
          item={formModal.item}
          woId={selectedWOId}
          workType={inferredWorkType}
          onClose={() => setFormModal(null)}
          onSaved={(savedItem) => {
            if (formModal.mode === 'add') {
              setMarkingItems(list => [...list, savedItem])
            } else {
              setMarkingItems(list =>
                list.map(i => i.item_id === savedItem.item_id ? savedItem : i))
            }
            setFormModal(null)
            setRowError('')
          }}
        />
      )}

      {/* Delete confirmation — single or bulk */}
      {deleteConfirm && (
        <ConfirmModal
          title={deleteConfirm.ids.length === 1
            ? 'Delete this marking item?'
            : `Delete ${deleteConfirm.ids.length} marking items?`}
          message={deleteConfirm.ids.length === 1
            ? `${deleteConfirm.label || 'Item'} will be permanently removed from the sheet.`
            : 'These rows will be permanently removed from the sheet. This cannot be undone.'}
          confirmLabel="Delete"
          cancelLabel="Cancel"
          danger
          onConfirm={() => doDelete(deleteConfirm.ids)}
          onCancel={() => setDeleteConfirm(null)}
        />
      )}

      <div>
        <h1 className="text-2xl font-black text-navy">Field Report</h1>
        <p className="text-slate-500 text-sm mt-0.5">Log crew sign-in, work progress, and photos for a Work Order.</p>
      </div>

      {formError && (
        <div ref={errorRef} className="bg-red-50 border border-red-200 text-red-700 rounded-xl p-4 text-sm whitespace-pre-line scroll-mt-24">{formError}</div>
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

            {rowError && (
              <div className="bg-red-50 border border-red-200 rounded-lg px-3 py-2 text-xs text-red-700">
                {rowError}
              </div>
            )}

            {!selectedWOId && (
              <p className="text-xs text-slate-400 italic">
                Select a Work Order above to see its marking items.
              </p>
            )}

            {selectedWOId && markingsLoading && (
              <p className="text-xs text-slate-400 italic">Loading marking items…</p>
            )}

            {selectedWOId && !markingsLoading && markingItems.length === 0 && (
              <p className="text-xs text-slate-400 italic bg-slate-50 rounded-lg px-3 py-2">
                No items pre-populated from the scan. Use "Add marking manually" below to log what was done.
              </p>
            )}

            {/* Bulk-select header — visible only while bulk mode is on. */}
            {bulkMode && (
              <div className="flex items-center justify-between bg-slate-50
                              border border-slate-200 rounded-lg px-3 py-2 sticky top-0 z-10">
                <label className="flex items-center gap-2 text-xs text-slate-600 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={selectedIds.size === markingItems.length && markingItems.length > 0}
                    onChange={e => e.target.checked ? selectAll() : clearSelection()}
                    className="w-4 h-4 accent-navy"
                  />
                  <span className="font-semibold">
                    {selectedIds.size} selected
                  </span>
                </label>
                <div className="flex gap-2 items-center">
                  <button type="button" onClick={exitBulkMode}
                    className="text-xs font-semibold text-slate-500 hover:text-slate-700 px-2">
                    Clear
                  </button>
                  <button type="button"
                    disabled={selectedIds.size === 0}
                    onClick={() => setDeleteConfirm({ ids: Array.from(selectedIds) })}
                    className="text-xs font-bold bg-red-500 text-white hover:bg-red-600
                               disabled:opacity-40 disabled:cursor-not-allowed
                               rounded-lg px-3 py-1.5 transition-colors">
                    Delete selected
                  </button>
                </div>
              </div>
            )}

            {/* Planned + manual items in sheet order, with a divider each
                time the Section value changes so the crew can read
                top-to-bottom against the paper WO. */}
            {markingItems.map((item, idx) => {
              const prevSection = idx > 0 ? markingItems[idx - 1].section : null
              const showHeader  = item.section !== prevSection
              return (
                <div key={item.item_id}>
                  {showHeader && (
                    <p className="text-[10px] font-bold uppercase tracking-wider text-slate-400 mt-3 mb-1">
                      {SECTION_HEADERS[item.section] || item.section || 'Other'}
                    </p>
                  )}
                  <MarkingItemRow
                    item={item}
                    selected={selectedIds.has(item.item_id)}
                    saving={rowSaving.has(item.item_id)}
                    bulkMode={bulkMode}
                    onToggleSelect={toggleSelect}
                    onLocalChange={onRowLocalChange}
                    onCommit={onRowCommit}
                    onEdit={(it) => setFormModal({ mode: 'edit', item: it })}
                    onDelete={(it) => setDeleteConfirm({
                      ids:   [it.item_id],
                      label: `${it.category}${it.intersection ? ` — ${it.intersection} ${it.direction}` : ''}`
                    })}
                    onStartBulk={enterBulkMode}
                  />
                </div>
              )
            })}

            {selectedWOId && (
              <button
                type="button"
                onClick={() => setFormModal({ mode: 'add', item: null })}
                className="btn-ghost text-xs mt-1"
              >
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
              <CrewCard key={i} idx={i} data={m} onChange={updateCrew} onRemove={removeCrew} workDate={workDate} />
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
