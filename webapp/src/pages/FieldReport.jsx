import { useState, useEffect, useMemo, useRef } from 'react'
import { useSearchParams } from 'react-router-dom'
import StatusBadge      from '../components/StatusBadge'
import ConfirmModal     from '../components/ConfirmModal'
import MarkingFormModal from '../components/MarkingFormModal'
import RowKebab         from '../components/RowKebab'
import {
  UNIT_OPTIONS, unitForCategory, unitIsLocked,
  pickLayout, rowIsCompletable, rowRequiresColor,
} from '../lib/markingCategories'
import { parseQty }    from '../lib/parseQty'
import { validateQty } from '../lib/qtyValidation'
import { opToday } from '../lib/dateOps'
import { usePhotoUploadQueue, reverseGeocode } from '../lib/photoUploadQueue'
import { formatWatermarkAddress } from '../lib/photoPipeline'

const SECTION_HEADERS = {
  'Top Table':          'WO Top Table',
  'Intersection Grid':  'Intersection Grid',
  'Manual':             'Manually Added',
}

// pickLayout / rowIsCompletable / unit helpers live in
// ../lib/markingCategories — single source of truth shared with the
// Add/Edit modal.

// ── Helpers ───────────────────────────────────────────────────
// Today's date as YYYY-MM-DD comes from `opToday()` in lib/dateOps —
// that's the operational-day-aware variant which buckets pre-cutoff
// (e.g. 02:30 AM) submissions back to yesterday's shift.

// ── HEIC → JPEG converter ─────────────────────────────────────
// iPhones ship HEIC by default. Browsers can't decode HEIC for
// <img> preview or canvas, so we convert to JPEG the moment a file
// is picked. The converted File is what flows through the rest of
// the pipeline (preview + watermark + compress + upload), so no other code
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

// Can the browser actually decode this file to pixels? HEIC that
// heic2any failed to convert (and any corrupt image) will fail here.
// We MUST catch it before queueing a stamped photo: once it reaches
// watermarkImage the failure is silent — it returns the original
// unstamped, and Drive ends up with an unstamped / undecodable file.
// Mirrors watermarkImage's own decode path (createImageBitmap, then an
// <img> fallback for the iOS-Safari cases createImageBitmap chokes on).
async function isDecodableImage(file) {
  try {
    const bmp = await createImageBitmap(file)
    bmp.close?.()
    return true
  } catch {
    return await new Promise(resolve => {
      const url = URL.createObjectURL(file)
      const img = new Image()
      img.onload  = () => { URL.revokeObjectURL(url); resolve(true) }
      img.onerror = () => { URL.revokeObjectURL(url); resolve(false) }
      img.src = url
    })
  }
}

// Image compression now lives inside usePhotoUploadQueue —
// compressed-on-upload (post-watermark for captures, raw for library).


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

// ── Work Order combobox ───────────────────────────────────────
// Button-trigger dropdown matching the CategorySelect / IntersectionSelect
// pattern from MarkingFormModal: the button shows the current selection
// (or a placeholder), clicking opens a panel with a dedicated search
// input above a scrollable list. No free-text on the trigger — picking
// from the list is the only way to set the WO.
function WOCombobox({ wos, selectedWOId, onSelect, disabled = false, disabledTitle = '' }) {
  const [open, setOpen]   = useState(false)
  const [query, setQuery] = useState('')
  const wrapRef   = useRef(null)
  const searchRef = useRef(null)

  const selected = wos.find(w => w.id === selectedWOId) || null
  const labelOf  = (wo) => `${wo.id} — ${wo.location} (${wo.borough})`

  // Sort WOs by the numeric portion of their ID (e.g. "RM-43282" → 43282)
  // so the dropdown reads in WO-number order regardless of how /api/wos
  // returned them. WOs whose ID has no digits sink to the end.
  const sortedWOs = useMemo(() => {
    return [...wos].sort((a, b) => {
      const na = parseInt(String(a.id || '').match(/\d+/)?.[0] ?? '', 10)
      const nb = parseInt(String(b.id || '').match(/\d+/)?.[0] ?? '', 10)
      if (Number.isNaN(na) && Number.isNaN(nb)) return String(a.id).localeCompare(String(b.id))
      if (Number.isNaN(na)) return 1
      if (Number.isNaN(nb)) return -1
      return na - nb
    })
  }, [wos])

  // Reset query + autofocus search whenever the panel opens; wire up
  // outside-click and Escape handlers so the panel feels native.
  useEffect(() => {
    if (!open) return
    setQuery('')
    setTimeout(() => searchRef.current?.focus(), 0)
    const onMouseDown = (e) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) setOpen(false)
    }
    const onKey = (e) => { if (e.key === 'Escape') setOpen(false) }
    document.addEventListener('mousedown', onMouseDown)
    document.addEventListener('keydown', onKey)
    return () => {
      document.removeEventListener('mousedown', onMouseDown)
      document.removeEventListener('keydown', onKey)
    }
  }, [open])

  const trimmed = query.trim().toLowerCase()
  const filtered = !trimmed
    ? sortedWOs
    : sortedWOs.filter(wo =>
        `${wo.id} ${wo.location || ''} ${wo.borough || ''}`
          .toLowerCase()
          .includes(trimmed)
      )

  // Force-close the panel any time the combobox becomes disabled (e.g.
  // admin entered completed-WO edit mode while the dropdown was open).
  useEffect(() => { if (disabled) setOpen(false) }, [disabled])

  return (
    <div ref={wrapRef} className="relative">
      <button
        type="button"
        onClick={() => { if (!disabled) setOpen(o => !o) }}
        disabled={disabled}
        title={disabled ? disabledTitle : undefined}
        className={`field-input w-full text-left flex items-center justify-between
                    ${disabled ? 'bg-slate-100 text-slate-400 cursor-not-allowed' : ''}`}
      >
        <span className={selected ? (disabled ? 'truncate' : 'text-slate-800 truncate') : 'text-slate-400'}>
          {selected ? labelOf(selected) : '— Choose or search Work Order —'}
        </span>
        <span className="text-slate-400 text-xs ml-2 flex-shrink-0">▾</span>
      </button>
      {open && !disabled && (
        <div className="absolute z-30 mt-1 left-0 right-0 bg-white rounded-xl shadow-lg border border-slate-200 overflow-hidden">
          <div className="p-2 border-b border-slate-100">
            <input
              ref={searchRef}
              type="text"
              value={query}
              onChange={e => setQuery(e.target.value)}
              placeholder="Search work orders…"
              className="field-input text-base sm:text-sm"
            />
          </div>
          <div className="max-h-[260px] overflow-y-auto">
            {filtered.length === 0 && (
              <p className="px-3 py-2 text-xs text-slate-400 italic">
                No matching work orders
              </p>
            )}
            {filtered.map(wo => {
              const isSelected = wo.id === selectedWOId
              return (
                <button
                  key={wo.id}
                  type="button"
                  onClick={() => { onSelect(wo.id); setOpen(false) }}
                  className={`w-full text-left px-3 py-2 text-sm hover:bg-slate-50
                              ${isSelected ? 'bg-navy/5 font-semibold text-navy' : 'text-slate-700'}`}
                >
                  {labelOf(wo)}
                </button>
              )
            })}
          </div>
        </div>
      )}
    </div>
  )
}

// ── Yes / No toggle ───────────────────────────────────────────
function YNToggle({ value, onChange, yesLabel='Yes', noLabel='No', disabled=false }) {
  return (
    <div className={`flex border border-slate-200 rounded-lg overflow-hidden
                     ${disabled ? 'opacity-60' : ''}`}>
      {[{v:'no',l:noLabel,cls:'bg-red-500 text-white'},{v:'yes',l:yesLabel,cls:'bg-green-600 text-white'}].map(({v,l,cls})=>(
        <button key={v} type="button"
          onClick={()=> { if (!disabled) onChange(v) }}
          disabled={disabled}
          className={`flex-1 py-2.5 text-sm font-semibold border-l first:border-l-0 border-slate-200 transition-all
            ${disabled ? 'cursor-not-allowed' : ''}
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
  forceUnlock = false,
}) {
  const layout = pickLayout(item)
  // forceUnlock = the admin Completed-WO edit mode is on; ignore the
  // per-row Completed lock so qty / direction / etc. can be edited
  // inline. saveField on the parent passes preserve_completion=true
  // while edit mode is on so this doesn't wipe Date Completed.
  const locked = item.status === 'Completed' && !forceUnlock

  // text-base on mobile (16px) so iOS Safari doesn't auto-zoom the
  // viewport when a Qty / intersection / direction input is focused;
  // sm: drops back to the compact 14px we use on desktop.
  const INPUT = 'field-input text-base sm:text-sm py-2'
  const RO    = INPUT + ' bg-slate-50 text-slate-500 cursor-default focus:ring-0'
  const LOCK  = INPUT + ' bg-slate-50 text-slate-500 cursor-not-allowed focus:ring-0'
  const QTY_LIVE   = INPUT + ' text-center'
  const QTY_LOCKED = LOCK  + ' text-center'

  // onChange handlers. Text inputs update local state only; commits
  // happen in onBlur/onKeyDown. Unit (dropdown) commits immediately.
  //
  // Quantity gets an extra char-filter so the field accepts only what
  // parseQty understands (digits, `.`, `+`, `*`, `x`, `X`, whitespace).
  // The full QWERTY keyboard on mobile makes every non-numeric character
  // typable; without this filter, paste or stray taps could land letters
  // / units / currency symbols in the field.
  const onLocalText  = (field) => (e) => {
    let v = e.target.value
    if (field === 'quantity') {
      v = (v.match(/[\d.+*xX\s]/g) || []).join('')
    }
    onLocalChange(item.item_id, field, v)
  }
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
  // type="text" + full QWERTY keyboard (no inputMode override) — gives
  // crews a familiar layout where every operator the Qty parser
  // supports is a labelled key on screen: digits, `.`, `+`, `*`, `x`,
  // `X` (no hidden iOS long-press gestures to discover). type="text"
  // suppresses the type=number stepper / wheel / arrow-key spinning.
  // The numeric-only enforcement that used to come from the tel
  // keypad's limited charset is now done by the per-field filter in
  // onLocalText above — pasted units / currency / typed letters are
  // stripped before they hit state.
  const QtyBox = (
    <input
      type="text" placeholder="Qty"
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

  // Default-layout rows with an intersection set get an extra read-only
  // intersection cell so "between" locations (e.g. "5 AV – 6 AV") on
  // top-table items like Lane Lines / Double Yellow Lines are visible
  // at a glance. When intersection is empty the row falls back to the
  // compact 3-cell layout.
  const hasInt = layout === 'default' && !!String(item.intersection || '').trim()

  // Grid templates — [checkbox?] + fields + [action]. Checkbox col is
  // 0px when bulk mode is off (hidden entirely).
  const CB  = bulkMode ? '24px ' : ''
  const END = ' 28px'
  // Desktop (sm+) — single-line, identical to the original layout.
  const tplDesktop = layout === 'grid'
    ? `${CB}1fr 110px 64px 90px 56px${END}`
    : layout === 'mma'
      ? `${CB}1fr 1fr 90px 56px${END}`
      : hasInt
        ? `${CB}1fr 110px 90px 56px${END}`
        : `${CB}1fr 90px 56px${END}`
  // Mobile (< sm) — for grid / mma / default-with-intersection layouts
  // the row stacks onto two lines: line 1 = [bulk-cb?] [type] [action],
  // line 2 = the contextual fields. Plain default (no intersection)
  // stays single-line — just slightly tighter.
  const tplMobileLine1   = `${CB}1fr ${END.trim()}`
  const tplMobileLine2Grid       = '1fr 56px 80px 48px'
  const tplMobileLine2Mma        = '1fr 80px 48px'
  const tplMobileLine2DefaultInt = '1fr 80px 48px'

  // Extract the intersection / direction cells so the same input
  // node can render in both the desktop and mobile branches.
  const IntersectionBox = (layout === 'grid' || hasInt) ? (
    <input type="text" readOnly value={item.intersection ?? ''}
      placeholder="Intersection" className={RO} />
  ) : null
  const DirectionBox = layout === 'grid' ? (
    <input type="text" readOnly value={item.direction ?? ''}
      placeholder="Direction" className={RO + ' text-center'} />
  ) : null

  const DescNote = item.description
    ? <p className={`text-[11px] text-slate-400 ${bulkMode ? 'pl-9' : 'pl-1'}`}>Note: {item.description}</p>
    : null

  const isMultiCell = layout === 'grid' || layout === 'mma' || hasInt

  return (
    <div className="space-y-0.5">
      {/* Desktop — single-line grid (unchanged from before). */}
      <div className="hidden sm:grid items-center gap-2"
           style={{ gridTemplateColumns: tplDesktop }}>
        {CheckboxCell}
        {CategoryBox}
        {layout === 'grid' && <>{IntersectionBox}{DirectionBox}</>}
        {layout === 'mma' && ColorBox}
        {layout === 'default' && hasInt && IntersectionBox}
        {QtyBox}
        {UnitBox}
        {ActionCell}
      </div>

      {/* Mobile — two lines for grid / mma / default-with-intersection,
          single line for plain default. */}
      <div className="sm:hidden">
        {isMultiCell ? (
          <div className="space-y-2">
            <div className="grid items-center gap-2"
                 style={{ gridTemplateColumns: tplMobileLine1 }}>
              {CheckboxCell}
              {CategoryBox}
              {ActionCell}
            </div>
            <div className="grid items-center gap-2"
                 style={{
                   gridTemplateColumns:
                     layout === 'grid' ? tplMobileLine2Grid
                     : layout === 'mma' ? tplMobileLine2Mma
                     : tplMobileLine2DefaultInt
                 }}>
              {layout === 'grid' && <>{IntersectionBox}{DirectionBox}</>}
              {layout === 'mma' && ColorBox}
              {layout === 'default' && hasInt && IntersectionBox}
              {QtyBox}
              {UnitBox}
            </div>
          </div>
        ) : (
          <div className="grid items-center gap-2"
               style={{ gridTemplateColumns: tplDesktop }}>
            {CheckboxCell}
            {CategoryBox}
            {QtyBox}
            {UnitBox}
            {ActionCell}
          </div>
        )}
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

// ── WO info panel ─────────────────────────────────────────────
// `due_date` arrives from Apps Script as a value the JS Date
// constructor can parse (ISO string, epoch ms, or already a Date).
// Render it as e.g. "May 25, 2026" — the raw `Date.toString()` slips
// out as "Mon May 25 2026 00:00:00 GMT-0400 (Eastern Daylight Time)"
// when the value reaches the field via JSON, which is unreadable.
function formatDueDate(raw) {
  if (!raw) return '—'
  const d = raw instanceof Date ? raw : new Date(raw)
  if (isNaN(d.getTime())) return String(raw)
  return new Intl.DateTimeFormat('en-US', {
    month: 'short', day: 'numeric', year: 'numeric',
    timeZone: 'America/New_York',
  }).format(d)
}

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
        ['Due Date',   formatDueDate(wo.due_date)],
        ['Work Type',  wo.work_type||'—'],
      ].map(([k,v])=>(
        <div key={k} className="flex gap-3">
          <span className="text-[11px] text-slate-400 font-semibold min-w-[60px] sm:min-w-[72px]">{k}</span>
          <span className="text-sm text-slate-700 font-medium">{v}</span>
        </div>
      ))}
      {/* Quick-access link to the WO's Drive folder — same destination
          as the 📁 icon on the WO Tracker tab. Hidden if the folder URL
          isn't cached on the row yet (Apps Script falls back to null). */}
      {wo.folder_url && (
        <div className="pt-1.5">
          <a
            href={wo.folder_url}
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center gap-1.5 text-xs font-bold text-navy
                       hover:underline"
          >
            View WO <span aria-hidden>📁</span>
          </a>
        </div>
      )}
    </div>
  )
}

// ── Photo capture + gallery ───────────────────────────────────
// Driven entirely by usePhotoUploadQueue. Two capture buttons:
//   Take Photo   — opens camera (capture="environment") AND fires
//                  navigator.geolocation in parallel. The captured file
//                  flows through HEIC convert → addCapture(file, geo) →
//                  watermark + Drive upload pipeline.
//   From Library — multi-select, accepts any image. HEIC converts.
//                  Skips watermark (timestamp/location would be wrong).
//
// Gallery renders every item the queue knows about: historic (already
// on Drive), in-flight (watermarking/uploading), and uploaded-this-
// session — sorted newest first. Delete on any item routes through a
// confirmation modal owned by the parent.
function PhotoCaptureGallery({ queue, woContext, onRequestDelete, onRequestPreview }) {
  const cameraRef  = useRef(null)
  const libraryRef = useRef(null)

  // Library-upload stamp modal. When the user picks photos from their
  // library we ask whether to burn on a date/time/location stamp (for
  // legacy photos taken before live capture stamped them). `stampModal`
  // holds the HEIC-normalised File[] awaiting that decision plus the
  // form fields.
  const [stampModal, setStampModal] = useState(null) // { files } | null
  const [stampDateTime, setStampDateTime] = useState('') // datetime-local string
  const [stampLat, setStampLat] = useState('')
  const [stampLng, setStampLng] = useState('')
  const [stampAddr, setStampAddr] = useState(null)   // string[] | null (looked-up preview)
  const [stampLooking, setStampLooking] = useState(false)
  const [stampError, setStampError] = useState('')

  const acquireGeo = () => new Promise(resolve => {
    if (!navigator.geolocation) { resolve(null); return }
    let settled = false
    const done = (v) => { if (!settled) { settled = true; resolve(v) } }
    navigator.geolocation.getCurrentPosition(
      pos => done({ lat: pos.coords.latitude, lng: pos.coords.longitude, accuracy: pos.coords.accuracy }),
      ()  => done(null),
      { enableHighAccuracy: true, timeout: 8000, maximumAge: 30000 }
    )
    setTimeout(() => done(null), 9000)
  })

  // iOS Safari hands the camera input a generic "image.jpg" filename for
  // every capture, which lands in Drive as image.jpg, image (1).jpg, etc.
  // Rename to {WO}_{LOCATION}_{YYYY-MM-DD}_{HH-MM-SS}.jpg before queueing
  // so the folder is readable and individual files are self-describing.
  const buildFilename = (d = new Date()) => {
    const woId = woContext?.id || 'WO'
    const loc  = (woContext?.location || '')
      .replace(/[^a-zA-Z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '')
      .slice(0, 40) || 'site'
    const pad = (n) => String(n).padStart(2, '0')
    const date = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`
    const time = `${pad(d.getHours())}-${pad(d.getMinutes())}-${pad(d.getSeconds())}`
    return `${woId}_${loc}_${date}_${time}.jpg`
  }
  // Materialise the rename as a fresh File. The original camera File's
  // `name` is read-only, so we have to wrap the bytes in a new one.
  const renameFile = (file, name) => new File([file], name, { type: file.type || 'image/jpeg' })

  const onCaptureChange = async (e) => {
    const file = e.target.files?.[0]
    e.target.value = ''
    if (!file) return
    const geo = await acquireGeo()
    const prepared = isHeic(file) ? await convertHeicToJpeg(file) : file
    const named = renameFile(prepared, buildFilename())
    queue.addCapture(named, geo)
  }

  const onLibraryChange = async (e) => {
    const files = Array.from(e.target.files || [])
    e.target.value = ''
    if (files.length === 0) return
    // Open the modal immediately in a "preparing" state. HEIC conversion
    // (heic2any is a 1.3 MB chunk + a multi-second decode) otherwise runs
    // with zero feedback and the button looks frozen.
    setStampDateTime('')
    setStampLat('')
    setStampLng('')
    setStampAddr(null)
    setStampError('')
    setStampModal({ files: null, skipped: [], preparing: true })
    // Convert any HEIC, then decode-check each file. Anything the browser
    // can't decode (a HEIC heic2any couldn't convert, a corrupt file) is
    // skipped rather than silently uploaded unstamped/undecodable.
    const ready = []
    const skipped = []
    for (const f of files) {
      const converted = isHeic(f) ? await convertHeicToJpeg(f) : f
      if (await isDecodableImage(converted)) ready.push(converted)
      else skipped.push(f.name || 'photo')
    }
    setStampModal({ files: ready, skipped, preparing: false })
  }

  // Name a batch of library files, suffixing _N when more than one is
  // picked at once so same-second captures don't collide in Drive.
  const nameBatch = (files, when) => files.map((f, i) => {
    const base = buildFilename(when).replace(/\.jpg$/, '')
    const suffix = files.length > 1 ? `_${i + 1}` : ''
    return renameFile(f, `${base}${suffix}.jpg`)
  })

  // "Just upload" — current behaviour, no stamp.
  const submitPlainUpload = () => {
    if (!stampModal?.files?.length) { setStampModal(null); return }
    queue.addLibrary(nameBatch(stampModal.files, new Date()))
    setStampModal(null)
  }

  // Reverse-geocode the entered coords so the user previews the exact
  // address lines that will render — same endpoint + formatter the live
  // capture pipeline uses, so the preview matches the burned-in result.
  const lookupAddress = async () => {
    const lat = parseFloat(stampLat), lng = parseFloat(stampLng)
    if (!isFinite(lat) || !isFinite(lng)) {
      setStampError('Enter a valid latitude and longitude first.')
      return
    }
    setStampError('')
    setStampLooking(true)
    try {
      const data = await reverseGeocode(lat, lng)
      setStampAddr(formatWatermarkAddress(data)) // [] when geocode unavailable
    } finally {
      setStampLooking(false)
    }
  }

  // "Add stamp & upload" — routes through the manual watermark path with
  // the hand-entered time + coords (+ looked-up address if previewed).
  const submitStampedUpload = async () => {
    if (!stampModal?.files?.length) { setStampError('No usable photos to upload.'); return }
    const lat = parseFloat(stampLat), lng = parseFloat(stampLng)
    if (!stampDateTime) { setStampError('Enter the date & time.'); return }
    if (!isFinite(lat) || !isFinite(lng)) {
      setStampError('Enter a valid latitude and longitude.')
      return
    }
    const when = new Date(stampDateTime) // datetime-local → device-local Date
    if (isNaN(when.getTime())) { setStampError('That date & time isn’t valid.'); return }
    // If the user never tapped "Look up address", resolve it now so the
    // address lines are deterministic rather than racing in the queue.
    let addressLines = stampAddr
    if (addressLines == null) {
      setStampLooking(true)
      try { addressLines = formatWatermarkAddress(await reverseGeocode(lat, lng)) }
      finally { setStampLooking(false) }
    }
    queue.addManual(nameBatch(stampModal.files, when), {
      captured_at: when.toISOString(),
      geo: { lat, lng },
      addressLines,
    })
    setStampModal(null)
  }

  return (
    <div className="space-y-3">
      <div className="grid grid-cols-2 gap-2">
        <button type="button"
          onClick={() => cameraRef.current?.click()}
          className="flex flex-col items-center justify-center gap-1 py-4
                     bg-navy text-white rounded-xl font-semibold text-sm
                     hover:bg-navy/90 active:bg-navy/80 transition-colors">
          <span className="text-2xl leading-none">📷</span>
          Take Photo
        </button>
        <button type="button"
          onClick={() => libraryRef.current?.click()}
          className="flex flex-col items-center justify-center gap-1 py-4
                     bg-white text-navy rounded-xl font-semibold text-sm
                     border-2 border-navy/15 hover:border-navy/40 active:bg-slate-50 transition-colors">
          <span className="text-2xl leading-none">🖼️</span>
          From Library
        </button>
        <input ref={cameraRef}  type="file" accept="image/*" capture="environment"
               className="hidden" onChange={onCaptureChange} />
        <input ref={libraryRef} type="file" accept="image/*" multiple
               className="hidden" onChange={onLibraryChange} />
      </div>
      <p className="text-[11px] text-slate-400">
        Live photos get a date/time/location stamp automatically. Uploading from
        your library? You'll be asked whether to stamp those too. All photos save
        to Drive automatically.
      </p>

      {queue.historicLoading && (
        <p className="text-[12px] text-slate-500 flex items-center gap-2">
          <span className="w-3 h-3 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
          Loading photos…
        </p>
      )}
      {queue.historicError && (
        <p className="text-[12px] text-amber-700">{queue.historicError}</p>
      )}

      {queue.items.length > 0 ? (
        <div className="flex flex-wrap gap-2">
          {queue.items.map(item => (
            <PhotoThumb key={item.id} item={item}
              onDelete={() => onRequestDelete(item)}
              onRetry={() => queue.retryOne(item.id)}
              onOpen={() => onRequestPreview(item)} />
          ))}
        </div>
      ) : !queue.historicLoading && (
        <p className="text-[11px] text-slate-400">No photos yet for this WO.</p>
      )}

      {(queue.uploadedCount > 0 || queue.pendingCount > 0 || queue.errorCount > 0) && (
        <p className="text-[11px] font-semibold flex flex-wrap gap-x-3 gap-y-1">
          {queue.uploadedCount > 0 && (
            <span className="text-green-700">✓ {queue.uploadedCount} on Drive</span>
          )}
          {queue.pendingCount > 0 && (
            <span className="text-slate-500">Uploading {queue.pendingCount}…</span>
          )}
          {queue.errorCount > 0 && (
            <span className="text-red-600">✕ {queue.errorCount} failed</span>
          )}
        </p>
      )}

      {/* Surfaces WHY any photo failed so the user isn't left guessing at
          the "✕ N failed" badge. Each failed thumbnail also has its own
          Retry button — this just makes the reason legible. */}
      <PhotoErrorNotice items={queue.items} />

      {/* Library-upload stamp prompt. Asks whether to burn a date/time/
          location overlay onto the picked photos before upload — for
          legacy photos taken before live capture stamped them. */}
      {stampModal && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center p-4"
          style={{ backgroundColor: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(2px)' }}
          onClick={() => setStampModal(null)}
        >
          <div
            className="bg-white rounded-2xl shadow-2xl w-full max-w-sm p-6 space-y-4 max-h-[90vh] overflow-y-auto"
            onClick={e => e.stopPropagation()}
          >
            {stampModal.preparing ? (
              <div className="py-8 flex flex-col items-center gap-3 text-center">
                <span className="w-8 h-8 border-4 border-slate-200 border-t-navy rounded-full animate-spin" />
                <p className="text-sm font-semibold text-slate-600">Preparing photos…</p>
                <p className="text-[11px] text-slate-400">
                  HEIC photos are converted in your browser, which can take a few seconds.
                </p>
              </div>
            ) : (
              <>
                <div className="text-center space-y-1">
                  <div className="text-3xl">🏷️</div>
                  <h2 className="text-lg font-black text-navy">Add a date &amp; location stamp?</h2>
                  {stampModal.files.length > 0 ? (
                    <p className="text-slate-500 text-sm leading-relaxed">
                      {stampModal.files.length === 1
                        ? 'This photo'
                        : `These ${stampModal.files.length} photos`} can be stamped with a
                      timestamp and geotag, matching live captures. Leave it off to
                      upload as-is.
                    </p>
                  ) : (
                    <p className="text-slate-500 text-sm leading-relaxed">
                      None of the selected photos could be read in this browser.
                    </p>
                  )}
                </div>

                {/* Files the browser couldn't decode (HEIC heic2any failed
                    on, corrupt images). Surfaced loudly instead of silently
                    uploading them unstamped / undecodable. */}
                {stampModal.skipped?.length > 0 && (
                  <div className="text-[11px] text-amber-800 bg-amber-50 border border-amber-200 rounded-lg p-2.5 space-y-1.5">
                    <p className="font-bold">
                      ⚠️ Couldn't read {stampModal.skipped.length} photo{stampModal.skipped.length > 1 ? 's' : ''} — likely
                      a HEIC this browser can't decode:
                    </p>
                    <ul className="list-disc list-inside">
                      {stampModal.skipped.map((n, i) => <li key={i} className="truncate">{n}</li>)}
                    </ul>
                    <p>
                      Convert {stampModal.skipped.length > 1 ? 'them' : 'it'} to JPEG and try again.
                      On a Mac: open in Preview → File → Export → Format: JPEG.
                    </p>
                  </div>
                )}

                {stampModal.files.length > 0 && (
                  <div className="space-y-3 border-t border-slate-100 pt-4">
                    <label className="block">
                      <span className="section-label">Date &amp; time</span>
                      <input
                        type="datetime-local" step="1"
                        value={stampDateTime}
                        onChange={e => setStampDateTime(e.target.value)}
                        className="field-input w-full"
                      />
                    </label>
                    <div className="grid grid-cols-2 gap-2">
                      <label className="block">
                        <span className="section-label">Latitude</span>
                        <input
                          type="text" inputMode="decimal" placeholder="40.750505"
                          value={stampLat}
                          onChange={e => { setStampLat(e.target.value); setStampAddr(null) }}
                          className="field-input w-full"
                        />
                      </label>
                      <label className="block">
                        <span className="section-label">Longitude</span>
                        <input
                          type="text" inputMode="decimal" placeholder="-73.999877"
                          value={stampLng}
                          onChange={e => { setStampLng(e.target.value); setStampAddr(null) }}
                          className="field-input w-full"
                        />
                      </label>
                    </div>

                    <button
                      type="button" onClick={lookupAddress} disabled={stampLooking}
                      className="text-xs font-bold text-navy hover:underline disabled:opacity-50"
                    >
                      {stampLooking ? 'Looking up…' : '🔎 Look up address'}
                    </button>
                    {stampAddr && (
                      <div className="text-[11px] text-slate-600 bg-slate-50 border border-slate-200 rounded-lg p-2 space-y-0.5">
                        {stampAddr.length
                          ? stampAddr.map((l, i) => <div key={i}>{l}</div>)
                          : <div className="italic text-slate-400">No address found — coordinates only will be shown.</div>}
                      </div>
                    )}
                  </div>
                )}
                {stampError && <p className="text-[11px] text-red-600">{stampError}</p>}

                <div className="flex flex-col gap-2 pt-1">
                  {stampModal.files.length > 0 && (
                    <>
                      <button
                        onClick={submitStampedUpload} disabled={stampLooking}
                        className="w-full py-3 rounded-xl font-bold text-sm bg-navy text-white
                                   hover:opacity-90 active:opacity-80 transition-all disabled:opacity-50"
                      >
                        Add stamp &amp; upload
                      </button>
                      <button
                        onClick={submitPlainUpload}
                        className="w-full py-3 rounded-xl font-bold text-sm bg-slate-100
                                   text-slate-600 hover:bg-slate-200 transition-all"
                      >
                        Upload without stamp
                      </button>
                    </>
                  )}
                  <button
                    onClick={() => setStampModal(null)}
                    className="w-full py-2 rounded-xl font-semibold text-xs text-slate-400
                               hover:text-slate-600 transition-all"
                  >
                    {stampModal.files.length > 0 ? 'Cancel' : 'Close'}
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}
    </div>
  )
}

function PhotoErrorNotice({ items }) {
  const failed = items.filter(i => i.status === 'error')
  if (failed.length === 0) return null
  // Collapse identical messages so three matching timeouts read as one.
  const reasons = [...new Set(failed.map(i => i.error || 'Upload failed'))]
  return (
    <div className="text-[11px] text-red-700 bg-red-50 border border-red-200
                    rounded-lg px-2.5 py-2 space-y-0.5">
      <p className="font-semibold">
        {failed.length} photo{failed.length === 1 ? '' : 's'} failed to upload to Drive
      </p>
      {reasons.map((r, i) => <p key={i} className="text-red-600">• {r}</p>)}
      <p className="text-red-500/80">
        Tap <span className="font-semibold">Retry</span> on the photo, or remove it and take it again.
      </p>
    </div>
  )
}

function PhotoThumb({ item, onDelete, onRetry, onOpen }) {
  const src = item.previewUrl
    || (item.thumbnail_b64 ? `data:${item.mime || 'image/jpeg'};base64,${item.thumbnail_b64}` : null)
  const inFlight = item.status === 'pending'
    || item.status === 'geocoding'
    || item.status === 'watermarking'
    || item.status === 'uploading'
  const canOpen = item.status === 'uploaded' || (!inFlight && src)
  return (
    <div className="relative group">
      <button type="button"
        onClick={canOpen ? onOpen : undefined}
        disabled={!canOpen}
        className={`block w-16 h-16 rounded-lg overflow-hidden border
          ${item.status === 'error' ? 'border-red-300' : 'border-slate-200'}
          ${canOpen ? 'cursor-zoom-in hover:ring-2 hover:ring-navy/40' : 'cursor-default'}
          ${inFlight ? 'opacity-60' : ''}`}>
        {src ? (
          <img src={src} alt={item.filename || 'photo'}
            className="w-full h-full object-cover" />
        ) : (
          <div className="w-full h-full bg-slate-50" />
        )}
      </button>
      {inFlight && (
        <div className="absolute inset-0 flex items-center justify-center pointer-events-none">
          <div className="w-5 h-5 border-2 border-white border-t-navy rounded-full animate-spin
                          bg-white/40 shadow" />
        </div>
      )}
      {item.status === 'uploaded' && item.source !== 'historic' && (
        <div className="absolute bottom-0.5 right-0.5 w-4 h-4 rounded-full bg-green-600
                        text-white text-[9px] flex items-center justify-center font-bold shadow pointer-events-none">✓</div>
      )}
      {item.status === 'error' && (
        <button type="button" onClick={(e) => { e.stopPropagation(); onRetry() }}
          title={item.error || 'Retry'}
          className="absolute inset-x-0 bottom-0 bg-red-600 text-white text-[9px]
                     font-semibold rounded-b-lg py-0.5 hover:bg-red-700">
          Retry
        </button>
      )}
      <button type="button" onClick={(e) => { e.stopPropagation(); onDelete() }}
        className="absolute -top-1.5 -right-1.5 w-5 h-5 bg-red-500 text-white rounded-full
                   text-xs font-bold flex items-center justify-center shadow
                   opacity-0 group-hover:opacity-100 focus:opacity-100 transition-opacity">×</button>
    </div>
  )
}

// Full-size lightbox. Routes the source two ways:
//   • Session items (previewUrl set)  → use the object URL → zero fetch.
//   • Historic items (drive_file_id)  → /api/wo-photos/:fileId/content
//     which proxies the Drive bytes back through Express.
// Click backdrop or Esc closes. Outer container handles the keyboard
// listener so the rest of the page doesn't get hijacked.
function PhotoLightbox({ item, onClose }) {
  const src = item.previewUrl
    ? item.previewUrl
    : item.drive_file_id
      ? `/api/wo-photos/${encodeURIComponent(item.drive_file_id)}/content`
      : null
  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onClose() }
    document.addEventListener('keydown', onKey)
    return () => document.removeEventListener('keydown', onKey)
  }, [onClose])
  const capturedLabel = item.captured_at
    ? new Date(item.captured_at).toLocaleString('en-US', {
        month: 'short', day: 'numeric', year: 'numeric',
        hour: 'numeric', minute: '2-digit',
      })
    : null
  return (
    <div className="fixed inset-0 z-50 flex flex-col items-center justify-center p-4"
      style={{ backgroundColor: 'rgba(0,0,0,0.85)' }}
      onClick={onClose}>
      <button type="button"
        onClick={onClose}
        className="absolute top-4 right-4 w-9 h-9 rounded-full bg-white/15 text-white
                   text-lg font-bold flex items-center justify-center hover:bg-white/25">×</button>
      {src ? (
        <img src={src} alt={item.filename || 'photo'}
          onClick={(e) => e.stopPropagation()}
          className="max-h-[85vh] max-w-[95vw] object-contain rounded shadow-2xl" />
      ) : (
        <div className="text-white text-sm">No preview available.</div>
      )}
      <div className="mt-3 text-white text-[12px] text-center space-y-0.5 px-2"
        onClick={(e) => e.stopPropagation()}>
        <div>{item.filename || 'photo'}</div>
        {capturedLabel && <div className="text-white/60">{capturedLabel}</div>}
        {item.drive_file_url && (
          <a href={item.drive_file_url} target="_blank" rel="noopener noreferrer"
            className="inline-block mt-1 text-white/80 hover:text-white underline">
            Open in Drive ↗
          </a>
        )}
      </div>
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
  const [workDate,      setWorkDate]       = useState(opToday())
  const [markingItems,  setMarkingItems]   = useState([])   // loaded from /api/wo-markings/:woId
  const [markingsLoading, setMarkingsLoading] = useState(false)
  const [crewChief,     setCrewChief]      = useState('')   // required — drives per-crew tagging
  const [employees,     setEmployees]      = useState([])   // Employee Registry, for crew chief picker
  const [issues,        setIssues]         = useState('')
  const [woComplete,    setWoComplete]     = useState('no')
  // Photos no longer live in component state — usePhotoUploadQueue owns
  // the IndexedDB-backed list. Captures + library imports stream to
  // Drive on add (not on submit), so submit-time work shrinks to a
  // "wait until queue.pendingCount === 0" check.
  const photoQueue = usePhotoUploadQueue(selectedWOId)
  const [photoDeleteConfirm, setPhotoDeleteConfirm] = useState(null)  // {item} or null
  const [photoLightbox, setPhotoLightbox] = useState(null)            // item or null
  // Completed WOs hide the photo uploader by default; this reveals it so
  // a crew can still add photos to a finished report (e.g. legacy photos
  // they forgot to attach). Reset whenever the selected WO changes.
  const [showCompletedUploader, setShowCompletedUploader] = useState(false)

  // Per-row UI state for the Marking Items list
  const [bulkMode,      setBulkMode]      = useState(false)
  const [selectedIds,   setSelectedIds]   = useState(new Set())
  const [rowSaving,     setRowSaving]     = useState(new Set())
  const [formModal,     setFormModal]     = useState(null)  // {mode, item} or null
  const [deleteConfirm, setDeleteConfirm] = useState(null)  // {ids:[...], label?} or null
  const [qtyConfirm,    setQtyConfirm]    = useState(null)  // {itemId, category, parsedStr, message} or null
  const [rowError,      setRowError]      = useState('')

  // Distinct intersections (in seeded order) + derived "X – Y" between
  // pairs, fed into the Add/Edit modal's Intersection combobox. Source
  // of truth = the WO's marking items themselves (intersection grid was
  // seeded from d.intersection_grid at WO intake).
  const woIntersections = useMemo(() => {
    const seen = new Set(), out = []
    markingItems.forEach(it => {
      if (it.section !== 'Intersection Grid') return
      const v = String(it.intersection || '').trim()
      if (!v || seen.has(v)) return
      seen.add(v); out.push(v)
    })
    return out
  }, [markingItems])
  const woBetweens = useMemo(() => {
    const out = []
    for (let i = 1; i < woIntersections.length; i++) {
      out.push(`${woIntersections[i - 1]} – ${woIntersections[i]}`)
    }
    return out
  }, [woIntersections])
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
  // "All marking items look complete — did you mean to mark the WO complete?"
  // suggestion modal. Fires when the user submits with the WO Complete
  // toggle still on No but every Marking Item is fully filled in.
  const [showCompleteSuggestModal, setShowCompleteSuggestModal] = useState(false)
  // Imperative override for the "Mark Complete & Submit" path so doSubmit
  // sees the flipped value without waiting for the setWoComplete state
  // update to flush. Read once + reset inside doSubmit.
  const completeOverrideRef = useRef(null)
  // Soft-warn modal when a Crew Chief is likely wrapping up last night's
  // shift but the form's date is set to today. shape: { priorDate, todayDate } | null
  const [shiftAttribModal, setShiftAttribModal] = useState(null)
  // Refs let doSubmit re-enter (after the modal closes) with the
  // user-picked date and skip the check loop. Both consume-once.
  const submitDateOverrideRef = useRef(null)
  const skipShiftAttribCheckRef = useRef(false)
  // Completed-WO edit mode. When the loaded WO is Completed, the page
  // opens read-only with a banner. Clicking Edit:
  //   - snapshots the current marking items into `originalItems`
  //   - flips editMode → true
  //   - unlocks marking-item rows (forceUnlock) + the bottom Update panel
  //
  // In edit mode every inline edit / modal save / kebab delete BUFFERS
  // client-side into the live `markingItems` state. NOTHING hits the
  // server until the admin clicks one of the three save mode buttons
  // (Save Data Only / Replace CFR / Save as New CFR). On Cancel — or
  // a beforeunload abandonment — the diff is discarded and the WO is
  // untouched on the server.
  const [editMode, setEditMode] = useState(false)
  const [originalItems, setOriginalItems] = useState(null)  // snapshot or null
  // Used to mint client-side IDs for items created via "+ Add manually"
  // in edit mode. Format: "_temp_<n>" — recognisable by both the diff
  // computer and the server-side batched add logic. Real IDs come back
  // after refetchMarkings post-save.
  const tempIdCounterRef = useRef(0)
  const nextTempId = () => `_temp_${++tempIdCounterRef.current}`

  // Cancel-confirm modal for "Discard N changes?" — null = closed,
  // otherwise { count }.
  const [editCancelConfirm, setEditCancelConfirm] = useState(null)

  // Build the { adds, edits, deletes } diff between the snapshot taken
  // on Edit click and the live markingItems state.
  //   adds:    items with _temp_ IDs (locally created)
  //   edits:   real-IDed items whose comparable fields differ from snapshot
  //   deletes: real-IDed items in snapshot but not in current
  // Newly-added items deleted within the session disappear entirely
  // (they never had a server-side row, so no DELETE call needed).
  const computeBufferDiff = () => {
    if (!originalItems) return { adds: [], edits: [], deletes: [] }
    const origMap = new Map(originalItems.map(i => [i.item_id, i]))
    const currMap = new Map(markingItems.map(i => [i.item_id, i]))
    const FIELDS = ['category','intersection','direction','description',
                    'quantity','unit','color_material','notes']
    const eq = (a, b) => String(a ?? '') === String(b ?? '')
    const adds = []
    const edits = []
    for (const item of markingItems) {
      if (String(item.item_id || '').startsWith('_temp_')) {
        adds.push(item)
        continue
      }
      const orig = origMap.get(item.item_id)
      if (!orig) continue   // server-created since snapshot — unusual; skip
      let changed = false
      const patch = { item_id: item.item_id }
      for (const f of FIELDS) {
        if (!eq(orig[f], item[f])) {
          changed = true
          patch[f] = item[f]
        }
      }
      if (changed) edits.push(patch)
    }
    const deletes = []
    for (const orig of originalItems) {
      if (String(orig.item_id || '').startsWith('_temp_')) continue
      if (!currMap.has(orig.item_id)) deletes.push(orig.item_id)
    }
    return { adds, edits, deletes }
  }
  const hasBufferChanges = () => {
    const d = computeBufferDiff()
    return d.adds.length + d.edits.length + d.deletes.length > 0
  }

  // Enter Edit mode — snapshot the current items + reset temp counter.
  const enterEditMode = () => {
    setOriginalItems(markingItems.map(i => ({ ...i })))
    tempIdCounterRef.current = 0
    setEditMode(true)
  }

  // Exit Edit mode — discard buffer (markingItems reverts to snapshot
  // if there's a snapshot; otherwise no-op).
  const discardEditBuffer = () => {
    if (originalItems) setMarkingItems(originalItems)
    setOriginalItems(null)
    setEditMode(false)
    setEditCancelConfirm(null)
  }

  // Cancel handler — short-circuit confirm when nothing changed.
  const onClickCancelEdit = () => {
    const d = computeBufferDiff()
    const n = d.adds.length + d.edits.length + d.deletes.length
    if (n === 0) {
      discardEditBuffer()
    } else {
      setEditCancelConfirm({ count: n })
    }
  }

  // beforeunload guard while there are unsaved buffered changes. The
  // browser shows its native "leave site?" dialog. We can't reliably
  // fire async reverts here — that's the whole point of the deferred-
  // PATCH design: there is no server state to revert. Just block the
  // accidental tab-close so the admin doesn't lose work silently.
  useEffect(() => {
    if (!editMode) return
    const handler = (e) => {
      if (!hasBufferChanges()) return
      e.preventDefault()
      e.returnValue = ''  // legacy browsers
      return ''
    }
    window.addEventListener('beforeunload', handler)
    return () => window.removeEventListener('beforeunload', handler)
  }, [editMode, originalItems, markingItems])  // eslint-disable-line react-hooks/exhaustive-deps
  // Modal driving the three regen modes (data_only / replace_cfr / new_cfr).
  // null = closed; otherwise { mode, state: 'submitting'|'success'|'error', ... }
  const [editCompletedSubmission, setEditCompletedSubmission] = useState(null)
  // Include-in-production checkbox + date. Defaults set per-mode at
  // click time; toggling persists within the session.
  const [includeInProduction, setIncludeInProduction] = useState(false)
  const [productionDate,      setProductionDate]      = useState(opToday())

  const selectedWO = wos.find(w => w.id === selectedWOId) ?? null
  // Completed WO is loaded — the page opens in read-only mode with an
  // Edit banner. Used to lock Date of Work, Issues/Notes, Add Marking,
  // Photo picker, and the Completion toggle until/while editing.
  const isCompletedWO = selectedWO && String(selectedWO.status).toLowerCase() === 'completed'
  // For fields that are editable in edit mode but locked otherwise.
  const lockedBecauseCompletedView = isCompletedWO && !editMode
  // For fields that stay locked even in edit mode (Date of Work, Photos,
  // Completion toggle — the edit-completed handler doesn't read them).
  const lockedAlwaysForCompleted   = isCompletedWO

  // ── MMA waterblasting gate ──────────────────────────────────────
  // Only MMA WOs (water_blast_required === 'Yes - MMA') need the
  // confirmation toggle. Thermo WOs stay "N/A" in col 13 and never
  // see the gate.  Toggle locks once confirmed — 3-dot menu → Edit
  // reopens the unlock path for the rare "oops, wasn't actually done"
  // case.
  const isMMAJob    = selectedWO?.water_blast_required === 'Yes - MMA'
  const wbConfirmed = selectedWO?.water_blast_confirmed === 'Yes'
  const wbGated     = isMMAJob && !wbConfirmed
  const [wbSubmitting,    setWbSubmitting]    = useState(false)
  const [wbConfirmModal,  setWbConfirmModal]  = useState(null)  // null | { nextValue: boolean }
  const [wbEditUnlocked,  setWbEditUnlocked]  = useState(false) // kebab → Edit unlocks the toggle on a confirmed WO

  // Reset the "unlock" state whenever the selected WO changes so the
  // locked pill is back by default on the next WO.
  useEffect(() => { setWbEditUnlocked(false) }, [selectedWOId])
  useEffect(() => { setShowCompletedUploader(false) }, [selectedWOId])
  // Reset edit mode + production-day defaults whenever the selected WO
  // changes. The completed-WO edit panel state is scoped per-WO.
  useEffect(() => {
    setEditMode(false)
    setOriginalItems(null)
    setIncludeInProduction(false)
    setProductionDate(opToday())
  }, [selectedWOId])

  const refreshWOs = async () => {
    const d = await fetch('/api/wos').then(r => r.json())
    if (d.error) throw new Error(d.error)
    setWOs(d.wos ?? [])
  }

  useEffect(() => {
    refreshWOs()
      .catch(e => setApiError(e.message))
      .finally(() => setLoading(false))
  }, [])

  // Load Employee Registry for the Crew Chief picker. Same endpoint
  // SignIn.jsx hits — server returns { employees: [{ name }, ...] }.
  // Best-effort; on failure the picker just renders empty and the user
  // gets blocked at submit with a clear "pick a Crew Chief" error.
  useEffect(() => {
    fetch('/api/employees')
      .then(r => r.json())
      .then(d => setEmployees(Array.isArray(d.employees) ? d.employees : []))
      .catch(err => console.warn('employees fetch failed:', err))
  }, [])

  // Deep-link from the Dashboard: /field-report?wo=RM-43282 pre-selects
  // the matching WO. Wait until wos are loaded so we can verify the id
  // exists — a stale link silently no-ops rather than locking the form
  // to a missing WO.
  const [searchParams] = useSearchParams()
  const deepLinkedWO   = searchParams.get('wo')
  useEffect(() => {
    if (!deepLinkedWO || selectedWOId) return
    if (wos.some(w => w.id === deepLinkedWO)) setSelectedWOId(deepLinkedWO)
  }, [deepLinkedWO, wos])  // eslint-disable-line react-hooks/exhaustive-deps

  async function doSetWaterblasting(nextValue) {
    if (!selectedWOId) return
    setWbSubmitting(true)
    try {
      const res = await fetch(`/api/waterblasting/${encodeURIComponent(selectedWOId)}/confirm`, {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ confirmed: !!nextValue }),
      })
      const data = await res.json().catch(() => ({}))
      if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
      await refreshWOs()
      // Collapse the unlock once the toggle has been set, so the kebab
      // pill is the only re-entry point for another edit.
      setWbEditUnlocked(false)
    } catch (err) {
      raiseError('Could not update waterblasting status: ' + err.message)
    } finally {
      setWbSubmitting(false)
    }
  }

  // Load pre-populated marking items from the WO scan whenever the
  // selected WO changes.  Important: clear the previous WO's items
  // *synchronously* before the fetch starts — otherwise stale rows
  // stay visible (and editable) under the loading indicator until the
  // new fetch returns, which caused crew to edit the wrong WO's
  // markings. Also guard against a stale fetch clobbering the list if
  // the user switches WOs again while the first request is still in
  // flight.
  useEffect(() => {
    setMarkingItems([])
    setSelectedIds(new Set())
    setRowSaving(new Set())
    inFlightRef.current.clear()

    if (!selectedWOId) {
      return
    }

    let cancelled = false
    setMarkingsLoading(true)
    fetch(`/api/wo-markings/${encodeURIComponent(selectedWOId)}`)
      .then(r => r.json())
      .then(d => {
        if (cancelled) return
        if (d.error) throw new Error(d.error)
        // Preserve sheet row order (= insertion order from the scan + any
        // earlier manual adds) — no client-side sorting needed.
        setMarkingItems(d.items || [])
      })
      .catch(e => {
        if (cancelled) return
        console.error('Failed to load marking items:', e)
        setMarkingItems([])
      })
      .finally(() => {
        if (!cancelled) setMarkingsLoading(false)
      })

    return () => { cancelled = true }
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
  //
  // Edit-completed mode buffers the change client-side instead of
  // PATCHing. The whole diff (adds / edits / deletes) flushes in one
  // server call when the admin clicks a save mode at the bottom of the
  // form. Cancel or beforeunload throws the buffer away — nothing was
  // ever persisted.
  const saveField = async (itemId, patch) => {
    if (editMode) {
      // Buffered path: just mutate local state. Validation that
      // belongs to the inline UI (e.g. numeric-only filtering) already
      // ran in onRowLocalChange before getting here.
      setMarkingItems(list => list.map(i => {
        if (i.item_id !== itemId) return i
        return { ...i, ...patch }
      }))
      return
    }
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
  // and on-change for dropdowns. For Qty: parse "15x10"-style
  // multiplication shorthand, then range-validate against the row's
  // Marking Type — out-of-range values open a confirmation modal
  // instead of saving directly.
  const onRowCommit = (itemId, field, value) => {
    if (field !== 'quantity') {
      saveField(itemId, { [field]: value })
      return
    }
    const parsedStr = parseQty(value)
    // If the multiplication shorthand resolved to a different string,
    // show the resolved number in the input immediately so what the
    // user sees matches what we'll save (and what the modal references).
    if (parsedStr !== value) {
      onRowLocalChange(itemId, 'quantity', parsedStr)
    }
    const item     = markingItems.find(i => i.item_id === itemId)
    const category = item?.category || ''
    const check    = validateQty(category, parsedStr)
    if (!check.ok) {
      setQtyConfirm({ itemId, category, parsedStr, message: check.message })
      return
    }
    const normalized = (parsedStr === '' || parsedStr == null)
      ? null
      : parseFloat(parsedStr)
    saveField(itemId, { quantity: normalized })
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

  // Wait for every queued photo to reach 'uploaded'. Reads the queue's
  // latest counts via the ref the hook exposes so we always see the
  // post-tick state (state-via-closure would be stale). Returns 0 on
  // clean drain, or the number of photos that flipped to 'error' so
  // the caller can surface a precise message.
  const photoQueueRef = useRef(photoQueue)
  useEffect(() => { photoQueueRef.current = photoQueue }, [photoQueue])
  const waitForPhotoQueue = async () => {
    // Soft cap: 5 minutes. Lets even a slow EDGE connection finish a
    // batch of 10 compressed photos; anything beyond that, the user is
    // better off cancelling and checking signal.
    const deadline = Date.now() + 5 * 60 * 1000
    while (Date.now() < deadline) {
      const q = photoQueueRef.current
      if (q.errorCount > 0) return q.errorCount
      if (q.pendingCount === 0) return 0
      await new Promise(r => setTimeout(r, 250))
    }
    return photoQueueRef.current.pendingCount
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
    // Edit-completed mode: buffer the delete client-side. If the item
    // is a newly-added temp item, it disappears completely. If it's
    // an existing-server-side item, it'll be DELETED in the batched
    // save call. Cancel/abandon restores everything from snapshot.
    if (editMode) {
      const set = new Set(ids)
      setMarkingItems(list => list.filter(i => !set.has(i.item_id)))
      setSelectedIds(prev => {
        const next = new Set(prev)
        ids.forEach(id => next.delete(id))
        if (bulkMode && next.size === 0) setBulkMode(false)
        return next
      })
      setDeleteConfirm(null)
      setRowError('')
      return
    }
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

  // Infer WO work type from any loaded planned item (Thermo beats MMA if
  // mixed). Falls back to blank if no items are loaded yet.
  const inferredWorkType = markingItems.some(i => String(i.work_type).toLowerCase() === 'thermo')
    ? 'Thermo'
    : markingItems.some(i => String(i.work_type).toLowerCase() === 'mma') ? 'MMA' : ''

  // ── Core submit (called after all guards pass) ────────────
  async function doSubmit() {
    setFormError('')
    setSubmitting(true)
    // "Mark Complete & Submit" sets the override so this submit posts
    // wo_complete=true even though setWoComplete('yes') won't have
    // flushed yet. One-shot — clear after read so a normal submit doesn't
    // accidentally inherit it.
    const completeOverride = completeOverrideRef.current
    completeOverrideRef.current = null
    const woCompleteFinal = completeOverride !== null ? completeOverride : (woComplete === 'yes')

    // Pre-flight: shift-attribution check. If the chief seems to be
    // wrapping up last night's overnight shift but workDate is today,
    // open a soft-warn modal and bail. The modal's pick handler sets
    // submitDateOverrideRef + skipShiftAttribCheckRef and re-calls
    // doSubmit, so the second pass uses the chosen date and skips the
    // check. Both refs are consumed on read.
    const dateOverride = submitDateOverrideRef.current
    submitDateOverrideRef.current = null
    const skipShiftCheck = skipShiftAttribCheckRef.current
    skipShiftAttribCheckRef.current = false
    const effDate = dateOverride || workDate

    if (!skipShiftCheck && crewChief.trim()) {
      try {
        const res = await fetch('/api/field-report/check-shift-attribution', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ crew_chief: crewChief.trim(), work_date: effDate }),
        })
        const data = await res.json().catch(() => ({}))
        if (data && data.should_confirm && data.prior_date) {
          setShiftAttribModal({ priorDate: data.prior_date, todayDate: effDate })
          setSubmitting(false)
          return
        }
      } catch (err) {
        // Network/server hiccup — fail open. Better to submit with the
        // user's intended date than to block them on a soft warning.
        console.warn('shift-attribution check failed:', err)
      }
    }

    // Step 0 — make sure any focused input flushes its save, then wait
    // for all in-flight PATCHes to land so finalize/rollup on the server
    // read the freshest Marking Items state.
    setSubmitStep('Saving pending edits…')
    if (typeof document !== 'undefined' && document.activeElement?.blur) {
      document.activeElement.blur()
    }
    await waitForSaves()

    // Step 1 — photos. Captures + library imports already stream to
    // Drive on add (see usePhotoUploadQueue), so submit just has to
    // wait for every item to reach status='uploaded'. Errored items
    // block submit entirely — the user has to retry or delete them.
    // The queue keeps draining in the background even while we wait,
    // so this stall is bounded by whatever network speed remains.
    if (photoQueue.errorCount > 0) {
      setSubmitStep('')
      raiseError(`${photoQueue.errorCount} photo${photoQueue.errorCount === 1 ? ' is' : 's are'} stuck uploading. Tap Retry on the thumbnail or remove them, then submit again.`)
      setSubmitting(false)
      return
    }
    if (photoQueue.pendingCount > 0) {
      setSubmitStep(`Waiting on ${photoQueue.pendingCount} photo${photoQueue.pendingCount === 1 ? '' : 's'} to finish uploading…`)
      const stalled = await waitForPhotoQueue()
      if (stalled) {
        setSubmitStep('')
        raiseError(`${stalled} photo${stalled === 1 ? ' is' : 's are'} stuck uploading. Tap Retry on the thumbnail or remove them, then submit again.`)
        setSubmitting(false)
        return
      }
    }
    const photosUploaded = photoQueue.uploadedCount > 0

    // Step 2 — submit field report
    setSubmitStep('Submitting report…')
    // Sign-in (crew + signatures) lives on its own tab now — the field
    // report only carries WO-level state + the Crew Chief identifier
    // that threads through WDL → SI queue → DSID → PL for per-crew
    // tracking on multi-crew shifts.
    const reportBody = {
      wo_id:       selectedWOId,
      date:        effDate,
      crew_chief:  crewChief.trim(),
      wo_complete: woCompleteFinal,
      work_type:   inferredWorkType,
      issues:          issues.trim(),
      photos_uploaded: photosUploaded,
    }

    try {
      // Marking Items are already live-persisted via per-row CRUD
      // endpoints. Submit only sends WO-level data.
      const res = await fetch('/api/field-report', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify(reportBody),
      })
      const data = await res.json()
      if (data.error) throw new Error(data.error)
      setSubmitted({ wo_id:data.wo_id, status:data.status, photos:photoQueue.uploadedCount })

      // Refresh the WO dropdown so a just-completed WO drops out without
      // the user having to hard-reload. Apps Script updates the Tracker
      // status synchronously on wo_complete=true, so /api/wos already
      // reflects the new state by the time we get here. Cheap GET.
      refreshWOs().catch(err => {
        console.warn('post-submit WO refresh failed:', err)
      })

      // Fire-and-forget: kick off CFR JSON generation on the server after
      // the user's seen the success screen. (Sign-in JSON is no longer
      // generated here — it's filed from the Sign-In tab.) Any failure
      // lands in Automation Log, not this UI.
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
    if (wbGated) {
      raiseError('Waterblasting must be confirmed before this MMA work order can be submitted. Toggle "Waterblasting Confirmed" to Yes at the top of the page.')
      return
    }
    if (!workDate)      { raiseError('Please enter the date of work.'); return }
    if (!crewChief.trim()) {
      raiseError('Please select a Crew Chief for this shift. Multiple crews on the same source job are tracked separately by chief.')
      return
    }
    // MMA items must always carry a Color/Material value, even on a
    // partial submit — the field is intrinsic to the item, not tied to
    // completion. (rowIsCompletable already enforces this when WO is
    // marked complete; this check covers partial submits too.)
    const missingColor = markingItems.filter(i =>
      rowRequiresColor(i) && !String(i.color_material || '').trim()
    )
    if (missingColor.length > 0) {
      const labels = missingColor.slice(0, 6).map(i => `• ${i.category}`).join('\n')
      const more = missingColor.length > 6 ? `\n…and ${missingColor.length - 6} more` : ''
      raiseError(
        `Color / Material is required for MMA items. ` +
        `${missingColor.length} item${missingColor.length===1?'':'s'} ` +
        `${missingColor.length===1?'is':'are'} missing it:\n\n${labels}${more}`
      )
      return
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
    // Soft prompt: WO toggle is still No but every Marking Item is fully
    // filled in. Likely the user forgot to flip the toggle; ask before
    // proceeding. Only triggers when items exist (an empty list isn't a
    // meaningful "all complete" state).
    if (
      woComplete === 'no'
      && markingItems.length > 0
      && markingItems.every(rowIsCompletable)
    ) {
      setShowCompleteSuggestModal(true)
      return
    }
    // Guard: no photos at all. Only fires when nothing is uploaded,
    // pending, OR errored — i.e. the user genuinely didn't add any.
    // If photos exist but failed (errorCount > 0) we fall through to
    // doSubmit, which surfaces the real "stuck uploading" message
    // instead of a misleading "did you forget photos?" prompt.
    if (photoQueue.uploadedCount === 0 && photoQueue.pendingCount === 0
        && photoQueue.errorCount === 0) {
      setShowPhotoModal(true); return
    }
    doSubmit()
  }

  // Helpers used by the "Did you mean to mark complete?" modal so the
  // photo guard still fires after the user resolves the suggestion.
  function continueAfterCompleteSuggest() {
    setShowCompleteSuggestModal(false)
    if (photoQueue.uploadedCount === 0 && photoQueue.pendingCount === 0
        && photoQueue.errorCount === 0) {
      setShowPhotoModal(true); return
    }
    doSubmit()
  }
  function markCompleteAndSubmit() {
    setWoComplete('yes')
    completeOverrideRef.current = true
    continueAfterCompleteSuggest()
  }

  // ── Reset ─────────────────────────────────────────────────
  function reset() {
    setSubmitted(null); setSelectedWOId(''); setWorkDate(opToday())
    setMarkingItems([]); setSelectedIds(new Set()); setRowSaving(new Set())
    setBulkMode(false); inFlightRef.current.clear()
    setCrewChief(''); setIssues(''); setRowError('')
    setWoComplete('no'); setFormError('')
    // photoQueue clears itself when selectedWOId flips back to '' above.
  }

  // ── Edit Completed WO submit ──────────────────────────────
  // Three modes: data_only (no CFR regen), replace_cfr (regen + replace
  // canonical WO doc), new_cfr (regen + archive as separate file with
  // _Updated_<asOfDate> suffix).
  async function submitEditCompleted(mode) {
    if (!selectedWOId) return
    // Flush any in-flight marking-item PATCHes before the server-side
    // rollup recompute — same pattern as the main submit (waitForSaves).
    if (typeof document !== 'undefined' && document.activeElement?.blur) {
      document.activeElement.blur()
    }
    setEditCompletedSubmission({ state: 'submitting', mode })
    await waitForSaves()
    try {
      const asOfDate = opToday()
      // Flush the deferred buffer of marking-item changes. Everything
      // the admin did in edit mode (inline edits, modal edits, manual
      // adds, kebab deletes) is encoded here as a single batched diff
      // — applied atomically(-ish) by the server in delete → edit →
      // add order before the Tracker write + CFR queue.
      const diff = computeBufferDiff()
      const body = {
        as_of_date:             asOfDate,
        issues:                 issues.trim(),
        regen_mode:             mode,
        include_in_production:  !!includeInProduction,
        production_date:        includeInProduction ? productionDate : '',
        marking_edits:          diff.edits,
        marking_adds:           diff.adds,
        marking_deletes:        diff.deletes,
      }
      const res = await fetch(`/api/wo/${encodeURIComponent(selectedWOId)}/edit-completed`, {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify(body),
      })
      const data = await res.json().catch(() => ({}))
      if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
      // Reconcile to server truth — picks up real item_ids for newly-
      // added rows, drops any temp-IDed items that the server didn't
      // accept, and covers any partial-failure scenarios.
      await refetchMarkings()
      setEditCompletedSubmission({ state: 'success', mode, result: data })
    } catch (err) {
      // Also reconcile on error so the user sees the actual current
      // state of the WO (some changes may have applied before the
      // failure) and can decide whether to retry.
      try { await refetchMarkings() } catch {}
      setEditCompletedSubmission({ state: 'error', mode, message: err.message || 'Update failed' })
    }
  }

  // ── Success ───────────────────────────────────────────────
  if (submitted) return (
    <div className="max-w-lg mx-auto px-4 py-16 text-center">
      <div className="text-6xl mb-5">✅</div>
      <h2 className="text-2xl font-black text-navy mb-2">Report Submitted!</h2>
      <p className="text-slate-500 mb-1">
        <strong className="font-mono text-navy">{submitted.wo_id}</strong>
        {' '}— queued for sign-in
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
    <div className="max-w-2xl mx-auto px-4 py-4 sm:py-6 space-y-4">

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

      {/* Full-size photo preview */}
      {photoLightbox && (
        <PhotoLightbox item={photoLightbox} onClose={() => setPhotoLightbox(null)} />
      )}

      {/* Photo delete confirmation — applies to both uploaded (Drive
          trash) and still-queued (IndexedDB only) photos. */}
      {photoDeleteConfirm && (
        <ConfirmModal
          title="Delete this photo?"
          message={photoDeleteConfirm.item.drive_file_id
            ? 'This will remove the photo from Drive. Field crews and admins will no longer see it on the WO.'
            : 'This photo hasn’t finished uploading. Removing it now will discard it.'}
          confirmLabel="Delete Photo"
          cancelLabel="Keep"
          danger
          onConfirm={async () => {
            const item = photoDeleteConfirm.item
            setPhotoDeleteConfirm(null)
            try {
              if (item.drive_file_id) {
                await photoQueue.deleteOne(item.id)
              } else {
                // In-flight cancel — doesn't hit Drive (nothing to trash)
                // and forcibly clears inFlightRef so a zombie kick can't
                // hold the slot.
                await photoQueue.cancelOne(item.id)
              }
            } catch (err) {
              raiseError('Couldn’t delete photo: ' + (err?.message || 'unknown error'))
            }
          }}
          onCancel={() => setPhotoDeleteConfirm(null)}
        />
      )}

      {/* "All marking items look complete — flip the WO Complete toggle?"
          suggestion modal. Three buttons because cancel ≠ submit-as-is ≠
          mark-complete-and-submit. */}
      {shiftAttribModal && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center p-4"
          style={{ backgroundColor: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(2px)' }}
          onClick={() => setShiftAttribModal(null)}
        >
          <div
            className="bg-white rounded-2xl shadow-2xl w-full max-w-sm p-6 space-y-4"
            onClick={e => e.stopPropagation()}
          >
            <div className="text-3xl text-center">🌙</div>
            <div className="text-center space-y-1.5">
              <h2 className="text-lg font-black text-navy">Is this for last night's shift?</h2>
              <p className="text-slate-500 text-sm leading-relaxed">
                <span className="font-semibold text-slate-700">{crewChief}</span> filed
                Field Reports dated{' '}
                <span className="font-mono text-slate-700">{shiftAttribModal.priorDate}</span>{' '}
                in the overnight window. This new report is set to{' '}
                <span className="font-mono text-slate-700">{shiftAttribModal.todayDate}</span>{' '}
                — pick the shift it actually covers.
              </p>
            </div>
            <div className="flex flex-col gap-2 pt-1">
              <button
                onClick={() => {
                  const priorDate = shiftAttribModal.priorDate
                  setShiftAttribModal(null)
                  setWorkDate(priorDate)
                  submitDateOverrideRef.current = priorDate
                  skipShiftAttribCheckRef.current = true
                  doSubmit()
                }}
                className="w-full py-3 rounded-xl font-bold text-sm bg-navy text-white
                           hover:opacity-90 active:opacity-80 transition-all"
              >
                Last night's shift ({shiftAttribModal.priorDate})
              </button>
              <button
                onClick={() => {
                  const todayDate = shiftAttribModal.todayDate
                  setShiftAttribModal(null)
                  submitDateOverrideRef.current = todayDate
                  skipShiftAttribCheckRef.current = true
                  doSubmit()
                }}
                className="w-full py-3 rounded-xl font-bold text-sm bg-amber-500 text-white
                           hover:bg-amber-600 active:opacity-80 transition-all"
              >
                Today's shift ({shiftAttribModal.todayDate})
              </button>
              <button
                onClick={() => setShiftAttribModal(null)}
                className="w-full py-3 rounded-xl font-bold text-sm bg-slate-100
                           text-slate-600 hover:bg-slate-200 transition-all"
              >
                Go Back
              </button>
            </div>
          </div>
        </div>
      )}

      {showCompleteSuggestModal && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center p-4"
          style={{ backgroundColor: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(2px)' }}
          onClick={() => setShowCompleteSuggestModal(false)}
        >
          <div
            className="bg-white rounded-2xl shadow-2xl w-full max-w-sm p-6 space-y-4"
            onClick={e => e.stopPropagation()}
          >
            <div className="text-3xl text-center">⚠️</div>
            <div className="text-center space-y-1.5">
              <h2 className="text-lg font-black text-navy">Mark Work Order Complete?</h2>
              <p className="text-slate-500 text-sm leading-relaxed">
                You've filled in every Marking Item for this WO but the
                "WO Complete" toggle is still set to No. Mark it complete
                before submitting?
              </p>
            </div>
            <div className="flex flex-col gap-2 pt-1">
              <button
                onClick={markCompleteAndSubmit}
                className="w-full py-3 rounded-xl font-bold text-sm bg-navy text-white
                           hover:opacity-90 active:opacity-80 transition-all"
              >
                Mark Complete & Submit
              </button>
              <button
                onClick={continueAfterCompleteSuggest}
                className="w-full py-3 rounded-xl font-bold text-sm bg-amber-500 text-white
                           hover:bg-amber-600 active:opacity-80 transition-all"
              >
                Submit Without Marking Complete
              </button>
              <button
                onClick={() => setShowCompleteSuggestModal(false)}
                className="w-full py-3 rounded-xl font-bold text-sm bg-slate-100
                           text-slate-600 hover:bg-slate-200 transition-all"
              >
                Go Back
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Marking Items — Add / Edit form modal */}
      {formModal && (
        <MarkingFormModal
          mode={formModal.mode}
          item={formModal.item}
          woId={selectedWOId}
          workType={inferredWorkType}
          wo_intersections={woIntersections}
          wo_betweens={woBetweens}
          deferred={editMode}
          tempIdFactory={nextTempId}
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
            <WOCombobox
              wos={wos}
              selectedWOId={selectedWOId}
              onSelect={setSelectedWOId}
              disabled={editMode}
              disabledTitle="Cancel or save the current edit before switching work orders."
            />
          </Field>
          {selectedWO && <WOPanel wo={selectedWO} />}
        </div>

        {/* Completed-WO banner — view mode + edit mode states.
            Shown only when the loaded WO has status === 'Completed'.
            View mode: explainer + [Edit] button. Edit mode: banner
            switches color + offers [Cancel]. */}
        {selectedWO && String(selectedWO.status).toLowerCase() === 'completed' && (
          editMode ? (
            <div className="rounded-xl border border-amber-300 bg-amber-50 p-3 flex items-start gap-3">
              <span className="text-xl leading-none">✎</span>
              <div className="flex-1 text-sm">
                <p className="font-bold text-amber-900">Editing completed WO</p>
                <p className="text-xs text-amber-800/80 mt-0.5">
                  Marking-item changes save inline. Pick a Save mode at
                  the bottom of the page to commit the update.
                </p>
              </div>
              <button
                type="button"
                onClick={onClickCancelEdit}
                className="text-xs font-bold px-3 py-1.5 rounded-lg
                           bg-white text-amber-800 border border-amber-300
                           hover:bg-amber-100"
              >
                Cancel
              </button>
            </div>
          ) : (
            <div className="rounded-xl border border-green-200 bg-green-50 p-3 flex items-start gap-3">
              <span className="text-xl leading-none">✓</span>
              <div className="flex-1 text-sm">
                <p className="font-bold text-green-900">This WO is completed</p>
                <p className="text-xs text-green-800/80 mt-0.5">
                  Marking items are read-only. Click Edit to fix a typo,
                  regenerate the CFR, or save an updated CFR (punch order).
                </p>
              </div>
              <button
                type="button"
                onClick={enterEditMode}
                className="text-xs font-bold px-3 py-1.5 rounded-lg
                           bg-navy text-white hover:opacity-90"
              >
                Edit
              </button>
            </div>
          )
        )}

        {/* 1b · Waterblasting Confirmation — MMA only */}
        {isMMAJob && (
          <WaterblastingCard
            confirmed={wbConfirmed}
            submitting={wbSubmitting}
            editUnlocked={wbEditUnlocked}
            onRequestToggle={(next) => setWbConfirmModal({ nextValue: next })}
            onUnlockEdit={() => setWbEditUnlocked(true)}
          />
        )}

        {/* Everything below the waterblasting card is gated for unconfirmed
            MMA WOs — greyed + pointer-events-none + submit blocked, until
            the toggle flips to Yes (either from the UI or the backend). */}
        <fieldset
          disabled={wbGated}
          className={wbGated ? 'opacity-40 pointer-events-none select-none' : ''}
        >
        <div className="space-y-4">

        {/* 2 · Work Details */}
        <div className="card p-4 space-y-4">
          <p className="section-label">Work Details</p>
          <Field label="Date of Work" required>
            <input
              type="date"
              value={workDate}
              onChange={e=>setWorkDate(e.target.value)}
              disabled={lockedAlwaysForCompleted}
              title={lockedAlwaysForCompleted
                ? 'Completed WO — use the "Include in production for" date in the Update panel below to log a new production day.'
                : undefined}
              className={`field-input ${lockedAlwaysForCompleted ? 'bg-slate-100 text-slate-400 cursor-not-allowed' : ''}`}
            />
          </Field>

          {/* Crew Chief — required. Drives per-crew tagging so multiple
              crews on the same source job get separate queue cards / PLs.
              Picker pulls from /api/employees (Employee Registry). */}
          <Field label="Crew Chief" required hint="The crew lead for this shift. Used to attribute hours, production, and downstream documents to this specific crew.">
            <select
              value={crewChief}
              onChange={e => setCrewChief(e.target.value)}
              disabled={lockedAlwaysForCompleted}
              className={`field-input ${lockedAlwaysForCompleted ? 'bg-slate-100 text-slate-400 cursor-not-allowed' : ''}`}
            >
              <option value="">— Select crew chief —</option>
              {employees.map(emp => (
                <option key={emp.name} value={emp.name}>{emp.name}</option>
              ))}
            </select>
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
                top-to-bottom against the paper WO. Inside the
                "Intersection Grid" section we add a second-level
                subheader that fires whenever the intersection changes,
                so the crew can scan vertically by intersection. */}
            {markingItems.map((item, idx) => {
              const prev = idx > 0 ? markingItems[idx - 1] : null
              const showSectionHeader = !prev || item.section !== prev.section
              const isInGrid = item.section === 'Intersection Grid'
              const intName = String(item.intersection || '').trim()
              const showIntHeader =
                isInGrid && intName &&
                (showSectionHeader || (prev && intName !== String(prev.intersection || '').trim()))
              return (
                <div key={item.item_id}>
                  {showSectionHeader && (
                    <p className="text-[10px] font-bold uppercase tracking-wider text-slate-400 mt-3 mb-1">
                      {SECTION_HEADERS[item.section] || item.section || 'Other'}
                    </p>
                  )}
                  {showIntHeader && (
                    <p className="text-[11px] font-bold uppercase tracking-wider text-slate-600 mt-2 mb-1 pl-2 border-l-2 border-slate-300">
                      {intName}
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
                    forceUnlock={editMode}
                  />
                </div>
              )
            })}

            {selectedWOId && !lockedBecauseCompletedView && (
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
            <textarea
              value={issues}
              onChange={e=>setIssues(e.target.value)}
              placeholder="Problems, delays, or anything admin should know…"
              rows={3}
              readOnly={lockedBecauseCompletedView}
              className={`field-input resize-none
                ${lockedBecauseCompletedView ? 'bg-slate-100 text-slate-500 cursor-not-allowed' : ''}`}
            />
          </Field>
        </div>

        {/* 3 · Photos */}
        <div className="card p-4 space-y-3">
          <p className="section-label">Site Photos</p>
          {!selectedWOId ? (
            <p className="text-sm text-slate-400 italic">Select a Work Order above to add photos.</p>
          ) : lockedAlwaysForCompleted && !showCompletedUploader ? (
            <div className="space-y-2">
              <p className="text-sm text-slate-500 italic">
                This Work Order is completed. You can still add photos to the
                report — useful for legacy photos that were taken on-site but
                never attached.
                {selectedWO?.folder_url
                  ? <> Or use the <a href={selectedWO.folder_url} target="_blank" rel="noopener noreferrer" className="font-bold text-navy hover:underline">View WO 📁</a> link to drop them into Drive directly.</>
                  : null}
              </p>
              <button
                type="button"
                onClick={() => setShowCompletedUploader(true)}
                className="w-full py-3 rounded-xl font-bold text-sm bg-white text-navy
                           border-2 border-navy/15 hover:border-navy/40 active:bg-slate-50 transition-colors"
              >
                ＋ Add photos to completed report
              </button>
            </div>
          ) : (
            <PhotoCaptureGallery
              queue={photoQueue}
              woContext={selectedWO ? { id: selectedWO.id, location: selectedWO.location } : null}
              onRequestDelete={(item) => setPhotoDeleteConfirm({ item })}
              onRequestPreview={(item) => setPhotoLightbox(item)}
            />
          )}
        </div>

        {/* 4 · Completion — sign-in (crew + signatures) is filed from
            the Sign-In tab once the night's work is wrapped up. */}
        <div className="card p-4 space-y-3">
          <p className="section-label">Completion</p>
          <Field label="Is this Work Order complete?">
            {lockedAlwaysForCompleted ? (
              <>
                <YNToggle value="yes" onChange={() => {}}
                  noLabel="No — more work needed" yesLabel="Yes — WO Done ✓"
                  disabled />
                <p className="text-[11px] text-slate-400 mt-1.5">
                  Locked — this WO is already Completed. Use the
                  Dashboard kebab → Change Status… to flip it back to
                  In Progress or set it to Returned.
                </p>
              </>
            ) : (
              <YNToggle value={woComplete} onChange={setWoComplete}
                noLabel="No — more work needed" yesLabel="Yes — WO Done ✓" />
            )}
          </Field>
        </div>

        </div>
        </fieldset>

        {/* Submit / Update panel — for completed WOs in edit mode, the
            three-mode update panel replaces the normal submit button.
            View mode on a completed WO hides everything (Edit button on
            the banner is the only call-to-action). Active WOs always
            show the regular Submit button. */}
        {selectedWO && String(selectedWO.status).toLowerCase() === 'completed' ? (
          editMode ? (
            <UpdateCompletedWOPanel
              includeInProduction={includeInProduction}
              setIncludeInProduction={setIncludeInProduction}
              productionDate={productionDate}
              setProductionDate={setProductionDate}
              submitting={!!editCompletedSubmission}
              onMode={(mode) => submitEditCompleted(mode)}
            />
          ) : null
        ) : (
          <button
            type="submit"
            disabled={submitting || wbGated}
            title={wbGated ? 'Waterblasting must be confirmed before submitting this MMA work order.' : undefined}
            className="btn-primary w-full text-base"
          >
            {submitting
              ? (submitStep || 'Submitting…')
              : wbGated
                ? 'Waterblasting not confirmed'
                : 'Submit Field Report'}
          </button>
        )}

        <div className="h-8" />
      </form>

      {/* Edit-completed-WO submission status modal */}
      {editCompletedSubmission && (
        <EditCompletedStatusModal
          state={editCompletedSubmission}
          onClose={() => {
            // Only allow close on success or error.
            if (editCompletedSubmission.state !== 'submitting') {
              if (editCompletedSubmission.state === 'success') {
                // Exit edit mode — buffer is now persisted server-side.
                setEditMode(false)
                setOriginalItems(null)
              }
              setEditCompletedSubmission(null)
            }
          }}
        />
      )}

      {/* "Discard N changes?" confirm — fires when admin clicks Cancel
          in edit mode and there's an unsaved buffer. */}
      {editCancelConfirm && (
        <ConfirmModal
          title={`Discard ${editCancelConfirm.count} unsaved change${editCancelConfirm.count === 1 ? '' : 's'}?`}
          message="Your in-progress edits to this completed WO will be thrown away. The WO and its marking items return to their original state."
          confirmLabel="Discard Changes"
          cancelLabel="Keep Editing"
          danger
          onConfirm={discardEditBuffer}
          onCancel={() => setEditCancelConfirm(null)}
        />
      )}

      {qtyConfirm && (
        <ConfirmModal
          title="Quantity outside typical range"
          message={qtyConfirm.message}
          confirmLabel="Yes, keep it"
          cancelLabel="Edit value"
          onConfirm={() => {
            const { itemId, parsedStr } = qtyConfirm
            const normalized = (parsedStr === '' || parsedStr == null)
              ? null
              : parseFloat(parsedStr)
            setQtyConfirm(null)
            saveField(itemId, { quantity: normalized })
          }}
          onCancel={() => {
            setQtyConfirm(null)
            // Revert to whatever the server still has saved for this row.
            refetchMarkings()
          }}
        />
      )}

      {wbConfirmModal && (
        <ConfirmModal
          title={wbConfirmModal.nextValue ? 'Confirm waterblasting complete?' : 'Mark waterblasting as NOT confirmed?'}
          message={wbConfirmModal.nextValue
            ? 'This Work Order requires waterblasting before MMA can be applied. Confirm it has been completed on-site before the crew starts the field report.'
            : 'The Field Report will be blocked until waterblasting is confirmed again. Continue?'}
          confirmLabel={wbConfirmModal.nextValue ? 'Yes, confirmed' : 'Yes, unconfirm'}
          danger={!wbConfirmModal.nextValue}
          onConfirm={() => {
            const next = wbConfirmModal.nextValue
            setWbConfirmModal(null)
            doSetWaterblasting(next)
          }}
          onCancel={() => setWbConfirmModal(null)}
        />
      )}
    </div>
  )
}

/**
 * WaterblastingCard — renders at the top of the Field Report whenever
 * the selected WO is MMA work. Drives the full-page gate: crew can't
 * submit until this flips to Yes.
 *
 * Visual states:
 *   - Unconfirmed: amber banner + unlocked Yes/No toggle, explains
 *     why the rest of the form is greyed out.
 *   - Confirmed (locked): green pill + 3-dot kebab → "Edit" unlocks
 *     the toggle so an accidental Yes can be corrected.
 *   - Confirmed (unlocked via Edit): same as Unconfirmed but starts
 *     at Yes and gives the crew a path back.
 */
function WaterblastingCard({ confirmed, submitting, editUnlocked, onRequestToggle, onUnlockEdit }) {
  const showToggle = !confirmed || editUnlocked

  return (
    <div className={`card p-4 space-y-3 border-2 ${confirmed ? 'border-green-200 bg-green-50/40' : 'border-amber-300 bg-amber-50/40'}`}>
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0 flex-1">
          <p className="section-label">Waterblasting Confirmation</p>
          <p className="text-xs text-slate-500 mt-0.5 leading-snug">
            This is an <span className="font-semibold">MMA</span> work order — waterblasting must be completed before the crew starts applying material.
          </p>
        </div>
        {confirmed && !editUnlocked && (
          <RowKebab
            items={[
              { label: 'Edit', onClick: onUnlockEdit },
            ]}
          />
        )}
      </div>

      {showToggle ? (
        <div className="space-y-2">
          <label className="field-label">Waterblasting Confirmed <span className="text-red-400 ml-0.5">*</span></label>
          <YNToggle
            value={confirmed ? 'yes' : 'no'}
            onChange={(next) => {
              if (submitting) return
              const nextBool = next === 'yes'
              if (nextBool === confirmed) return
              onRequestToggle(nextBool)
            }}
            noLabel="No — not yet"
            yesLabel="Yes — confirmed ✓"
          />
          {submitting && (
            <p className="text-xs text-slate-500 italic">Saving…</p>
          )}
          {!confirmed && (
            <p className="text-xs text-amber-700 bg-amber-100/60 border border-amber-200 rounded-lg px-3 py-2 leading-snug">
              The rest of this Field Report is locked until waterblasting is confirmed. Toggle to <span className="font-semibold">Yes</span> once it has been completed on-site.
            </p>
          )}
        </div>
      ) : (
        <div className="flex items-center gap-2">
          <span className="inline-flex items-center gap-1.5 bg-green-100 text-green-800
                           text-xs font-bold px-3 py-1.5 rounded-full">
            <svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor"><path d="M6.5 11.5L3 8l1.4-1.4L6.5 8.7l5.1-5.1L13 5l-6.5 6.5z"/></svg>
            Confirmed
          </span>
          <span className="text-xs text-slate-500">Use the ⋮ menu to edit.</span>
        </div>
      )}
    </div>
  )
}


// ── Update-Completed-WO panel ────────────────────────────────
// Three-button footer that replaces the normal Submit Field Report
// button when the loaded WO is Completed and the admin has entered
// edit mode. Each button maps to a regen_mode the backend understands.
// The "Include in production" checkbox + date picker control whether
// a Work Day Log row is appended for sign-in / Production Log coverage.
function UpdateCompletedWOPanel({
  includeInProduction, setIncludeInProduction,
  productionDate, setProductionDate,
  submitting, onMode,
}) {
  // Smart default the checkbox per-mode when the user picks one — but
  // only if they haven't manually touched it. Keeps the Save-as-New
  // flow ergonomic while letting power-users override.
  const click = (mode) => {
    if (!submitting) onMode(mode)
  }

  return (
    <div className="card p-4 space-y-3">
      <p className="section-label">Update Completed WO</p>

      <label className="flex items-center gap-2 cursor-pointer">
        <input
          type="checkbox"
          checked={includeInProduction}
          onChange={e => setIncludeInProduction(e.target.checked)}
          className="w-4 h-4 accent-navy"
        />
        <span className="text-sm text-slate-700">Include this work in production for</span>
        <input
          type="date"
          value={productionDate}
          onChange={e => setProductionDate(e.target.value)}
          disabled={!includeInProduction}
          className="field-input py-1 px-2 max-w-[160px]"
        />
      </label>
      <p className="text-[11px] text-slate-400 -mt-1.5 pl-6">
        Checked: appends a Work Day Log row so the work counts in
        Sign-In + Production Log queues for that day. Leave unchecked
        for typo / data-only fixes.
      </p>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 pt-1">
        <button
          type="button"
          disabled={submitting}
          onClick={() => click('data_only')}
          className="px-3 py-3 rounded-xl text-sm font-bold border
                     bg-white text-slate-700 border-slate-200
                     hover:border-navy/40 hover:bg-slate-50
                     disabled:opacity-40 disabled:cursor-not-allowed
                     leading-tight"
        >
          Save Data Only
          <span className="block text-[10px] font-normal text-slate-400 mt-0.5">no CFR regen</span>
        </button>
        <button
          type="button"
          disabled={submitting}
          onClick={() => click('replace_cfr')}
          className="px-3 py-3 rounded-xl text-sm font-bold border
                     bg-navy/90 text-white border-navy
                     hover:opacity-90
                     disabled:opacity-40 disabled:cursor-not-allowed
                     leading-tight"
        >
          Replace CFR
          <span className="block text-[10px] font-normal opacity-80 mt-0.5">regen &amp; replace</span>
        </button>
        <button
          type="button"
          disabled={submitting}
          onClick={() => {
            // Default-on the production checkbox for Save-as-New
            // (punch-order rework is genuinely new work). User can
            // still uncheck before clicking again.
            if (!includeInProduction) setIncludeInProduction(true)
            click('new_cfr')
          }}
          className="px-3 py-3 rounded-xl text-sm font-bold border
                     bg-amber-500 text-white border-amber-500
                     hover:bg-amber-600
                     disabled:opacity-40 disabled:cursor-not-allowed
                     leading-tight"
        >
          Save as New CFR
          <span className="block text-[10px] font-normal opacity-90 mt-0.5">punch-order rework</span>
        </button>
      </div>
    </div>
  )
}


// ── Edit-Completed submission status modal ───────────────────
// Blocking modal that surfaces the result of an edit-completed submit.
// `state.state`: 'submitting' | 'success' | 'error'. Backdrop click +
// Escape are disabled while submitting.
function EditCompletedStatusModal({ state, onClose }) {
  if (!state) return null
  const isSubmitting = state.state === 'submitting'
  const isSuccess    = state.state === 'success'
  const isError      = state.state === 'error'

  const titleByMode = {
    data_only:   'Saving data update…',
    replace_cfr: 'Regenerating CFR (replace)…',
    new_cfr:     'Generating new CFR (save as new)…',
  }
  const successByMode = {
    data_only:   'Data update saved. CFR not regenerated.',
    replace_cfr: 'New CFR queued — will replace the existing archived PDF after the next worker run + admin approval.',
    new_cfr:     'New CFR queued with "_Updated" suffix — saved alongside the original after the next worker run + admin approval.',
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4"
      style={{ backgroundColor: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(2px)' }}
      onClick={() => !isSubmitting && onClose()}
    >
      <div
        className="bg-white rounded-2xl shadow-2xl w-full max-w-md p-6 space-y-4"
        onClick={e => e.stopPropagation()}
      >
        {isSubmitting && (
          <>
            <div className="w-10 h-10 mx-auto border-[3px] border-slate-200 border-t-navy rounded-full animate-spin" />
            <p className="text-center text-sm font-bold text-navy">{titleByMode[state.mode] || 'Submitting…'}</p>
            <p className="text-center text-xs text-slate-500">Don't navigate away.</p>
          </>
        )}
        {isSuccess && (
          <>
            <div className="text-4xl text-center">✅</div>
            <p className="text-center text-base font-bold text-navy">Done</p>
            <p className="text-center text-sm text-slate-600">{successByMode[state.mode]}</p>
            {state.result?.wdl_appended && (
              <p className="text-center text-xs text-slate-500">Production-day row appended.</p>
            )}
            <button
              type="button"
              onClick={onClose}
              className="w-full py-3 rounded-xl font-bold text-sm bg-navy text-white hover:opacity-90"
            >
              Close
            </button>
          </>
        )}
        {isError && (
          <>
            <div className="text-4xl text-center">⚠️</div>
            <p className="text-center text-base font-bold text-red-700">Update failed</p>
            <p className="text-center text-sm text-slate-600 break-words">{state.message}</p>
            <button
              type="button"
              onClick={onClose}
              className="w-full py-3 rounded-xl font-bold text-sm bg-slate-100 text-slate-700 hover:bg-slate-200"
            >
              Close
            </button>
          </>
        )}
      </div>
    </div>
  )
}
