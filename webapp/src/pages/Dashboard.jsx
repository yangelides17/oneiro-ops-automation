import { useState, useEffect, useMemo, useRef } from 'react'
import { Link, useSearchParams } from 'react-router-dom'
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell
} from 'recharts'
import StatusBadge from '../components/StatusBadge'
import { opToday } from '../lib/dateOps'
import RevenueTab from './RevenueTab'
import ProductionTab from './ProductionTab'
import DocStatusTab from './DocStatusTab'
import DownloadDocumentsModal from '../components/DownloadDocumentsModal'
import DocStatusChips from '../components/DocStatusChips'
import GenerateDocModal from '../components/GenerateDocModal'
import StatusPickerModal from '../components/StatusPickerModal'
import DeleteWOModal from '../components/DeleteWOModal'
import RowKebab from '../components/RowKebab'
import { usePendingCounts } from '../lib/PendingCountsContext'
import { NavBadge } from '../App'

// ── API ───────────────────────────────────────────────────────
async function fetchDashboard() {
  const res = await fetch('/api/dashboard')
  if (!res.ok) throw new Error('HTTP ' + res.status)
  const data = await res.json()
  if (data.error) throw new Error(data.error)
  return data
}

// ── Date helpers ──────────────────────────────────────────────
// "Today" everywhere in the dashboard means the *operational* day —
// pre-cutoff (default 4 AM) submissions bucket back to yesterday's
// shift, matching how the Apps Script side writes Daily Sign-In Data.
const isoToday = () => opToday()
// Monday on-or-before the given local-date ISO string. Sun → 6 days
// back; Mon → 0; otherwise day-1 back.  Operates on parsed local Date
// so DST transitions don't skew by an hour.
function mondayOnOrBefore(isoYmd) {
  const [y, m, d] = isoYmd.split('-').map(Number)
  const dt = new Date(y, m-1, d)
  const dow = dt.getDay()   // 0=Sun, 1=Mon, ..., 6=Sat
  const delta = dow === 0 ? -6 : 1 - dow
  dt.setDate(dt.getDate() + delta)
  return [dt.getFullYear(), String(dt.getMonth()+1).padStart(2,'0'), String(dt.getDate()).padStart(2,'0')].join('-')
}
const thisWeekMondayIso = () => mondayOnOrBefore(isoToday())
const lastWeekMondayIso = () => {
  const [y, m, d] = thisWeekMondayIso().split('-').map(Number)
  const dt = new Date(y, m-1, d); dt.setDate(dt.getDate() - 7)
  return [dt.getFullYear(), String(dt.getMonth()+1).padStart(2,'0'), String(dt.getDate()).padStart(2,'0')].join('-')
}

// Turn a "YYYY-MM-DD" into "M/D/YYYY" for user-facing messages.
const prettyDate = (iso) => {
  if (!iso) return ''
  const [y, m, d] = iso.split('-').map(Number)
  return `${m}/${d}/${y}`
}

// ── Tools dropdown ────────────────────────────────────────────
// Lives next to "Refresh" in the Dashboard header.  Surfaces the
// same operations that used to live in the spreadsheet's custom menu
// (generate daily docs, generate certified payroll) so we don't
// depend on the standalone-script onOpen trigger, which has been
// flaky to install.
function ToolsMenu() {
  const [open,     setOpen]     = useState(false)
  const [picker,   setPicker]   = useState(null)   // null | 'daily' | 'cert'
  const [pickerVal,setPickerVal]= useState('')
  // Active GenerateDocModal action — { title, description, onConfirm } or null.
  const [modalAction, setModalAction] = useState(null)
  const wrapRef = useRef(null)

  // Close on outside-click / Esc (matches RowKebab pattern).
  useEffect(() => {
    if (!open && !picker) return
    const h = (e) => {
      if (!wrapRef.current) return
      if (wrapRef.current.contains(e.target)) return
      setOpen(false); setPicker(null)
    }
    const esc = (e) => {
      if (e.key === 'Escape') { setOpen(false); setPicker(null) }
    }
    document.addEventListener('mousedown', h)
    document.addEventListener('keydown', esc)
    return () => {
      document.removeEventListener('mousedown', h)
      document.removeEventListener('keydown', esc)
    }
  }, [open, picker])

  function runDailyDocs(isoDate) {
    setOpen(false); setPicker(null)
    const when = isoDate ? prettyDate(isoDate) : `today (${prettyDate(isoToday())})`
    setModalAction({
      title: 'Generate Daily Documents',
      description: (
        <span>
          <span className="block font-semibold text-slate-700">{when}</span>
          <span className="block mt-1">Sign-In Log, Production Log, and Contractor Field Reports for every WO with crew activity on this date.</span>
        </span>
      ),
      onConfirm: async () => {
        const res = await fetch('/api/tools/daily-documents', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ date: isoDate || '' }),
        })
        const data = await res.json().catch(() => ({}))
        if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
        const n = data.entries_found ?? 0
        return {
          message: n > 0
            ? `Generated for ${when} — ${n} sign-in entr${n === 1 ? 'y' : 'ies'} processed. Find filled docs in "Docs Needing Review" → Approvals tab.`
            : `No sign-in entries found for ${when}. Nothing generated.`,
        }
      },
    })
  }

  function runProcessApproved() {
    setOpen(false); setPicker(null)
    setModalAction({
      title: 'Process Approved Documents',
      description: (
        <span>
          Run the same archive cron that fires automatically every 10 min — emails any pending doc and files them into Drive. Safe to run alongside the cron (script-wide lock).
        </span>
      ),
      onConfirm: async () => {
        const res = await fetch('/api/tools/process-approved-docs', { method: 'POST' })
        const data = await res.json().catch(() => ({}))
        if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
        const archived = data.archived ?? 0
        const errored  = data.errored ?? 0
        if (data.skipped) {
          return { message: 'Another run is already in progress — it will finish shortly. No action taken here.' }
        }
        if (archived === 0 && errored === 0) {
          return { message: 'No pending docs in Approved Docs folder — nothing to process.' }
        }
        const parts = []
        if (archived > 0) parts.push(`archived ${archived}`)
        if (errored > 0)  parts.push(`${errored} failed → Archive Errors`)
        const msg = `Processed approved docs: ${parts.join(', ')}.`
        if (errored > 0) throw new Error(msg)
        return { message: msg }
      },
    })
  }

  function runCertPayroll(isoWeekStart) {
    setOpen(false); setPicker(null)
    setModalAction({
      title: 'Generate Certified Payroll',
      description: (
        <span>
          <span className="block font-semibold text-slate-700">Week of {prettyDate(isoWeekStart)}</span>
          <span className="block mt-1">One Certified Payroll JSON per contract+borough that had sign-in activity in this Monday–Sunday week.</span>
        </span>
      ),
      onConfirm: async () => {
        const res = await fetch('/api/tools/certified-payroll', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ week_start: isoWeekStart }),
        })
        const data = await res.json().catch(() => ({}))
        if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
        const n = data.contract_groups ?? 0
        return {
          message: n > 0
            ? `Generated for week of ${prettyDate(isoWeekStart)} — ${n} contract group${n === 1 ? '' : 's'}. Find filled docs in "Docs Needing Review → Certified Payroll" → Approvals tab.`
            : `No sign-in entries found for week of ${prettyDate(isoWeekStart)}. Nothing generated.`,
        }
      },
    })
  }

  function openPicker(kind) {
    setPicker(kind)
    setOpen(false)
    // Sensible defaults for each picker.
    setPickerVal(kind === 'daily' ? isoToday() : thisWeekMondayIso())
  }

  function submitPicker() {
    if (!pickerVal) return
    if (picker === 'daily') runDailyDocs(pickerVal)
    else if (picker === 'cert') {
      // Snap the chosen date to the Monday of that week so the
      // generator always receives a valid week-start.
      runCertPayroll(mondayOnOrBefore(pickerVal))
    }
  }

  return (
    <div ref={wrapRef} className="relative">
      <button
        type="button"
        onClick={() => setOpen(v => !v)}
        className="btn-ghost flex items-center gap-1.5"
      >
        <span>🛠️</span>
        Tools
        <span className="text-xs opacity-60">▾</span>
      </button>

      {open && (
        <div className="absolute right-0 mt-1 z-30 min-w-[280px]
                        bg-white rounded-lg shadow-lg border border-slate-200
                        py-1 text-sm overflow-hidden">
          <p className="px-3 pt-2 pb-1 text-[10px] font-bold uppercase tracking-wider text-slate-400">
            Generate Daily Documents
          </p>
          <button className="tools-item" onClick={() => runDailyDocs('')}>
            Today ({prettyDate(isoToday())})
          </button>
          <button className="tools-item" onClick={() => openPicker('daily')}>
            Custom date…
          </button>

          <div className="my-1 border-t border-slate-100" />

          <p className="px-3 pt-1 pb-1 text-[10px] font-bold uppercase tracking-wider text-slate-400">
            Generate Certified Payroll
          </p>
          <button className="tools-item" onClick={() => runCertPayroll(thisWeekMondayIso())}>
            This week (week of {prettyDate(thisWeekMondayIso())})
          </button>
          <button className="tools-item" onClick={() => runCertPayroll(lastWeekMondayIso())}>
            Last week (week of {prettyDate(lastWeekMondayIso())})
          </button>
          <button className="tools-item" onClick={() => openPicker('cert')}>
            Custom week…
          </button>

          <div className="my-1 border-t border-slate-100" />

          <p className="px-3 pt-1 pb-1 text-[10px] font-bold uppercase tracking-wider text-slate-400">
            Approved Docs
          </p>
          <button className="tools-item" onClick={runProcessApproved}>
            Process approved docs now
          </button>
        </div>
      )}

      {picker && (
        <div
          className="fixed inset-0 z-40 flex items-center justify-center p-4"
          style={{ backgroundColor: 'rgba(0,0,0,0.5)' }}
          onClick={() => setPicker(null)}
        >
          <div
            className="bg-white rounded-2xl shadow-2xl w-full max-w-sm p-5 space-y-4"
            onClick={e => e.stopPropagation()}
          >
            <h2 className="text-lg font-black text-navy">
              {picker === 'daily' ? 'Pick a date' : 'Pick any date in the target week'}
            </h2>
            <p className="text-xs text-slate-500 leading-snug">
              {picker === 'daily'
                ? 'Generates Sign-In Log, Production Log, and Contractor Field Reports for every work order with crew activity on this date.'
                : `We'll run certified payroll for the Monday-Sunday week containing the date you pick. Chosen date snaps to its week's Monday automatically.`}
            </p>
            <input
              type="date"
              value={pickerVal}
              onChange={e => setPickerVal(e.target.value)}
              className="field-input w-full"
            />
            {picker === 'cert' && pickerVal && (
              <p className="text-[11px] text-slate-500">
                Will run for week of <span className="font-mono">{prettyDate(mondayOnOrBefore(pickerVal))}</span>
              </p>
            )}
            <div className="flex gap-2 pt-1">
              <button
                type="button"
                onClick={() => setPicker(null)}
                className="flex-1 py-2.5 rounded-xl text-sm font-bold bg-slate-100 text-slate-600 hover:bg-slate-200"
              >
                Cancel
              </button>
              <button
                type="button"
                disabled={!pickerVal}
                onClick={submitPicker}
                className="flex-1 py-2.5 rounded-xl text-sm font-bold bg-navy text-white hover:opacity-90 disabled:opacity-50"
              >
                Generate
              </button>
            </div>
          </div>
        </div>
      )}

      <GenerateDocModal
        open={!!modalAction}
        title={modalAction?.title || ''}
        description={modalAction?.description || null}
        onConfirm={modalAction?.onConfirm || (async () => ({}))}
        onClose={() => setModalAction(null)}
      />
    </div>
  )
}

// ── Stat Card ─────────────────────────────────────────────────
function StatCard({ label, value, sub, color }) {
  return (
    <div className="card p-4 flex flex-col gap-1">
      <span className="text-[11px] font-bold uppercase tracking-widest text-slate-500">
        {label}
      </span>
      <span className={`text-3xl font-black leading-none ${color ?? 'text-navy'}`}>
        {value ?? '—'}
      </span>
      {sub && <span className="text-xs text-slate-400">{sub}</span>}
    </div>
  )
}

// ── Filter bar ────────────────────────────────────────────────
//
// Multi-select pills: clicking a pill toggles it in the selected Set,
// and clicking "All" clears the Set. Empty Set ≡ no filter (matches
// every WO for that category). Pills combine within a category with
// OR ("Received" + "In Progress" → either status passes) and across
// categories with AND (status AND contractor AND borough).
const ALL = 'All'

function FilterBar({ label, options, selected, onToggle, onClear }) {
  const allActive = selected.size === 0
  const pillClass = (active) =>
    `text-xs px-3 py-1 rounded-full border font-medium transition-all
     ${active
       ? 'bg-navy text-white border-navy'
       : 'bg-white text-slate-600 border-slate-200 hover:border-navy/40'
     }`
  return (
    <div className="flex items-center gap-2 flex-wrap">
      <span className="text-xs font-semibold text-slate-500">{label}</span>
      <button
        key={ALL}
        onClick={onClear}
        className={pillClass(allActive)}
      >
        {ALL}
      </button>
      {options.map(opt => (
        <button
          key={opt}
          onClick={() => onToggle(opt)}
          className={pillClass(selected.has(opt))}
        >
          {opt}
        </button>
      ))}
    </div>
  )
}

// ── WO Table row ──────────────────────────────────────────────
function WORow({ wo, flagged, onDocsChange, onChangeStatus, onDeleteWO }) {
  const isOverdue = wo.due_date && new Date(wo.due_date) < new Date()
    && wo.status.toLowerCase() !== 'completed'

  // Quantity cell: e.g. "1500 SF" or "240 LF". Empty / zero → em-dash.
  const qtyNum = parseFloat(wo.quantity)
  const qtyText = (!wo.quantity || isNaN(qtyNum) || qtyNum === 0)
    ? '—'
    : `${wo.quantity} ${wo.quantity_unit || ''}`.trim()

  // Progress cell: thin bar + "x / y" label. 0 / 0 → em-dash.
  const total = wo.markings_total ?? 0
  const done  = wo.markings_completed ?? 0
  const pct   = total > 0 ? Math.round((done / total) * 100) : 0

  return (
    <tr className={`border-b border-slate-100 text-sm transition-colors
                    hover:bg-slate-50 ${flagged ? 'bg-amber-50/60' : ''}`}>
      <td className="py-2.5 px-3 font-mono font-semibold whitespace-nowrap">
        <Link
          to={`/field-report?wo=${encodeURIComponent(wo.id)}`}
          className="text-navy hover:underline"
        >
          {wo.id}
        </Link>
        {flagged && (
          <span className="ml-1.5 text-amber-500" title="Needs attention">⚠</span>
        )}
      </td>
      <td className="py-2.5 px-3 text-slate-700">{wo.contractor || '—'}</td>
      <td className="py-2.5 px-3">
        <span className="bg-slate-100 text-slate-600 rounded px-1.5 py-0.5 text-[11px] font-semibold">
          {wo.borough}
        </span>
      </td>
      <td className="py-2.5 px-3 text-slate-600 max-w-[180px] truncate">
        {wo.location || '—'}
      </td>
      <td className={`py-2.5 px-3 whitespace-nowrap text-xs font-medium
                      ${isOverdue ? 'text-red-600 font-bold' : 'text-slate-500'}`}>
        {prettyDate(wo.due_date) || '—'}
        {isOverdue && ' ⚡'}
      </td>
      <td className="py-2.5 px-3"><StatusBadge status={wo.status} /></td>
      <td className="py-2.5 px-3 text-slate-600 text-xs whitespace-nowrap">
        {qtyText}
      </td>
      <td className="py-2.5 px-3 min-w-[120px]">
        {total === 0 ? (
          <span className="text-slate-300 text-xs">—</span>
        ) : (
          <div className="flex items-center gap-2">
            <div className="flex-1 h-1.5 bg-slate-200 rounded-full overflow-hidden">
              <div
                className="h-full bg-navy"
                style={{ width: `${pct}%` }}
              />
            </div>
            <span className="text-[11px] text-slate-500 whitespace-nowrap">
              {done} / {total}
            </span>
          </div>
        )}
      </td>
      <td className="py-2.5 px-3">
        <DocStatusChips
          woId={wo.id}
          docs={wo.docs}
          onChange={onDocsChange}
        />
      </td>
      <td className="py-2.5 px-3">
        {wo.folder_url ? (
          <a
            href={wo.folder_url}
            target="_blank"
            rel="noopener noreferrer"
            title="Open WO folder in Google Drive"
            className="text-xs font-bold px-2 py-1 rounded-lg
                       bg-slate-100 text-slate-700 hover:bg-slate-200"
          >
            📁
          </a>
        ) : (
          <span className="text-slate-300 text-xs px-2 py-1" title="Folder not yet created">—</span>
        )}
      </td>
      <td className="py-2.5 px-3 text-right">
        <RowKebab items={[
          { label: 'Change Status…', onClick: () => onChangeStatus?.(wo) },
          { label: 'Delete WO…',     onClick: () => onDeleteWO?.(wo), danger: true },
        ]} />
      </td>
    </tr>
  )
}

// ── Main ──────────────────────────────────────────────────────
// Tab id 'operations' is kept internally for back-compat with any
// existing `?tab=operations` URLs; the user-facing label is WO Tracker.
const TABS = [
  { id: 'operations', label: 'WO Tracker' },
  { id: 'doc_status', label: 'Doc Status' },
  { id: 'production', label: 'Production' },
  { id: 'revenue',    label: 'Revenue'    },
]

function TabStrip({ active, onChange }) {
  const { counts } = usePendingCounts()
  // Only the Doc Status tab carries a badge today — surfaces the
  // combined SI + PL + CP pending count after the user has visited
  // the tab once (cold-start endpoint intentionally skips it).
  const badgeFor = (id) => id === 'doc_status' ? counts.doc_status_pending : null
  return (
    <div className="flex items-center gap-1 border-b border-slate-200">
      {TABS.map(t => (
        <button
          key={t.id}
          onClick={() => onChange(t.id)}
          className={`text-sm font-semibold px-4 py-2 -mb-px border-b-2 transition-colors
            ${active === t.id
              ? 'border-navy text-navy'
              : 'border-transparent text-slate-500 hover:text-slate-700'
            }`}
        >
          {t.label}
          <NavBadge n={badgeFor(t.id)} />
        </button>
      ))}
    </div>
  )
}

export default function Dashboard() {
  const [searchParams, setSearchParams] = useSearchParams()
  // Default tab is "operations" — anything outside the known tab IDs
  // collapses back to operations so a typo'd URL doesn't break.
  const tabParam = searchParams.get('tab')
  const activeTab = TABS.some(t => t.id === tabParam) ? tabParam : 'operations'
  const setActiveTab = (id) => {
    const next = new URLSearchParams(searchParams)
    if (id === 'operations') next.delete('tab')
    else next.set('tab', id)
    setSearchParams(next, { replace: true })
  }

  const [data,       setData]       = useState(null)
  const [loading,    setLoading]    = useState(true)
  const [error,      setError]      = useState(null)
  // Multi-select filters: each is a Set<string> of currently-selected
  // pill values. Empty Set = no filter applied (the "All" pill is the
  // visual active state for empty).
  const [statusFilt, setStatusFilt] = useState(() => new Set())
  const [contFilt,   setContFilt]   = useState(() => new Set())
  const [boroughFilt,setBoroughFilt]= useState(() => new Set())
  const [search,     setSearch]     = useState('')
  const [lastRefresh,setLastRefresh]= useState(null)
  const [showDownloadModal, setShowDownloadModal] = useState(false)
  // WO row kebab actions — `null` = closed; otherwise the WO object.
  const [statusPickerWO, setStatusPickerWO] = useState(null)
  const [deleteWO,       setDeleteWO]       = useState(null)
  // Pagination: how many WO rows to render at once. Mostly an
  // initial-render perf knob — the API still returns the full list in
  // one shot, so backend latency is unchanged. Reduces DOM nodes by
  // ~10× on a typical Tracker, which keeps the table feeling snappy.
  const WO_PAGE_SIZE = 20
  const [visibleCount, setVisibleCount] = useState(WO_PAGE_SIZE)

  // Toggle / clear helpers for the multi-select filter Sets.
  // toggleX adds opt if missing, removes if present. clearX empties.
  // A new Set is always returned so React sees a fresh reference and
  // useMemo dependents recompute.
  const makeToggle = (setter) => (opt) => setter(prev => {
    const next = new Set(prev)
    if (next.has(opt)) next.delete(opt)
    else              next.add(opt)
    return next
  })
  const toggleStatus  = makeToggle(setStatusFilt)
  const toggleCont    = makeToggle(setContFilt)
  const toggleBorough = makeToggle(setBoroughFilt)
  const clearStatus   = () => setStatusFilt(new Set())
  const clearCont     = () => setContFilt(new Set())
  const clearBorough  = () => setBoroughFilt(new Set())

  const load = async () => {
    setLoading(true)
    setError(null)
    try {
      const d = await fetchDashboard()
      setData(d)
      setLastRefresh(new Date())
    } catch (e) {
      setError(e.message)
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => { load() }, [])

  // Doc-flag toggle handler — used by the DocStatusChips popover on
  // each WO row. Optimistically updates the WO's docs object in
  // local state, fires the network call, returns true on success so
  // the chip component can keep its optimistic update (or revert).
  // The Apps Script set_docs_sent action accepts both 'CFR' and
  // 'Field Report' as friendly names; we pass whichever the chip
  // sends through.
  const onDocsChange = async (woId, friendlyDocType, partial) => {
    // Map friendly doc_type → docs object key for the local update.
    const docKey = ({
      'CFR':               'cfr',
      'Field Report':      'cfr',
      'Production Log':    'production_log',
      'Sign-In':           'signin',
      'Certified Payroll': 'certified_payroll',
      'Invoice':           'invoice',
    })[friendlyDocType]

    if (docKey) {
      setData(prev => {
        if (!prev || !Array.isArray(prev.wos)) return prev
        const next = {
          ...prev,
          wos: prev.wos.map(w => w.id === woId
            ? { ...w, docs: { ...w.docs, [docKey]: { ...w.docs?.[docKey], ...partial } } }
            : w
          ),
        }
        return next
      })
    }
    try {
      const res = await fetch('/api/documents/flags', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ updates: [{ wo_id: woId, doc_type: friendlyDocType, ...partial }] }),
      })
      const body = await res.json().catch(() => ({}))
      if (!res.ok || body.error) {
        console.warn('onDocsChange failed:', body.error || `HTTP ${res.status}`)
        return false
      }
      return true
    } catch (e) {
      console.warn('onDocsChange threw:', e.message)
      return false
    }
  }

  // "Completed Today" is a pseudo-status — the pill adds one extra
  // filter beyond the real WO Tracker statuses (Received / Dispatched /
  // In Progress / Completed). Admin uses it at end-of-day to sanity-
  // check that every WO they expected to close has its docs in the
  // Approval queue.
  const COMPLETED_TODAY = 'Completed Today'

  // Compare a WO's Work End Date (col 18 via API `work_end`) to
  // *today in local time*. Cell can come back as either a Date-like
  // string ("Mon Apr 20 2026 …") or an ISO "YYYY-MM-DD" depending on
  // how Sheets stored it, so handle both.
  function isCompletedTodayLocal(wo) {
    if (String(wo.status).toLowerCase() !== 'completed') return false
    const raw = String(wo.work_end || '')
    if (!raw) return false
    const todayLocal = isoToday()
    const isoMatch = raw.match(/^(\d{4}-\d{2}-\d{2})/)
    if (isoMatch) return isoMatch[1] === todayLocal
    const d = new Date(raw)
    if (isNaN(d.getTime())) return false
    const y = d.getFullYear(), m = String(d.getMonth()+1).padStart(2,'0'), dd = String(d.getDate()).padStart(2,'0')
    return `${y}-${m}-${dd}` === todayLocal
  }

  // Filter + search WOs. Within a category: empty Set = no filter; a
  // populated Set means the WO must match one of its values (OR).
  // Across categories: AND. The Completed Today pseudo-status is a
  // calendar predicate, not a real status string, so it has its own
  // matcher inside the status OR-loop.
  const filteredWOs = useMemo(() => {
    if (!data?.wos) return []
    return data.wos.filter(wo => {
      if (statusFilt.size > 0) {
        let ok = false
        for (const s of statusFilt) {
          if (s === COMPLETED_TODAY) {
            if (isCompletedTodayLocal(wo)) { ok = true; break }
          } else if (wo.status === s) {
            ok = true; break
          }
        }
        if (!ok) return false
      }
      if (contFilt.size    > 0 && !contFilt.has(wo.contractor))   return false
      if (boroughFilt.size > 0 && !boroughFilt.has(wo.borough))   return false
      if (search) {
        const q = search.toLowerCase()
        return (
          wo.id.toLowerCase().includes(q) ||
          wo.location.toLowerCase().includes(q) ||
          wo.contractor.toLowerCase().includes(q)
        )
      }
      return true
    })
  }, [data, statusFilt, contFilt, boroughFilt, search])

  // Unique filter options
  const contractors = useMemo(() =>
    [...new Set((data?.wos ?? []).map(w => w.contractor).filter(Boolean))].sort()
  , [data])

  const boroughs = useMemo(() =>
    [...new Set((data?.wos ?? []).map(w => w.borough).filter(Boolean))].sort()
  , [data])

  // Bar chart data: status breakdown
  const chartData = data ? [
    { name: 'Received',    value: data.stats.received,    fill: '#3b82f6' },
    { name: 'Dispatched',  value: data.stats.dispatched,  fill: '#f59e0b' },
    { name: 'In Progress', value: data.stats.in_progress, fill: '#f97316' },
    { name: 'Completed',   value: data.stats.complete,    fill: '#22c55e' },
  ] : []

  // Loading state
  if (loading && !data) {
    return (
      <div className="flex items-center justify-center min-h-[50vh]">
        <div className="flex flex-col items-center gap-4">
          <div className="w-10 h-10 border-[3px] border-slate-200 border-t-navy
                          rounded-full animate-spin" />
          <p className="text-slate-400 text-sm">Loading dashboard…</p>
        </div>
      </div>
    )
  }

  if (error && !data) {
    return (
      <div className="flex items-center justify-center min-h-[50vh]">
        <div className="text-center max-w-sm">
          <div className="text-4xl mb-4">⚠️</div>
          <h2 className="text-lg font-bold text-slate-800 mb-2">Failed to load</h2>
          <p className="text-slate-500 text-sm mb-4">{error}</p>
          <button onClick={load} className="btn-outline text-sm px-4 py-2">
            Try again
          </button>
        </div>
      </div>
    )
  }

  const { stats, attention = [] } = data ?? {}

  return (
    <div className="max-w-6xl mx-auto px-4 py-6 space-y-6">

      {/* ── Page header ──────────────────────────────── */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-black text-navy leading-none">
            {activeTab === 'revenue'    ? 'Revenue Dashboard'
              : activeTab === 'production' ? 'Production Dashboard'
              : activeTab === 'doc_status' ? 'Doc Status'
              : 'Work Order Tracker'}
          </h1>
          {activeTab === 'operations' && lastRefresh && (
            <p className="text-slate-400 text-xs mt-1">
              Updated {lastRefresh.toLocaleTimeString()}
            </p>
          )}
        </div>
        <div className="flex items-center gap-2">
          <ToolsMenu />
          <button
            onClick={() => setShowDownloadModal(true)}
            className="btn-ghost flex items-center gap-1.5"
            title="Download a zip of archived docs filtered by mode/contractor/doc-type"
          >
            <span>📦</span>
            Documents
          </button>
          {activeTab === 'operations' && (
            <button
              onClick={load}
              disabled={loading}
              className="btn-ghost flex items-center gap-1.5"
            >
              <span className={loading ? 'animate-spin inline-block' : ''}>↻</span>
              Refresh
            </button>
          )}
          <Link to="/field-report" className="btn-primary text-sm px-4 py-2">
            + Field Report
          </Link>
        </div>
      </div>

      {/* ── Tab strip ────────────────────────────────── */}
      <TabStrip active={activeTab} onChange={setActiveTab} />

      {activeTab === 'revenue' ? (
        <RevenueTab />
      ) : activeTab === 'production' ? (
        <ProductionTab />
      ) : activeTab === 'doc_status' ? (
        <DocStatusTab />
      ) : (
        <OperationsTabContent
          stats={stats}
          attention={attention}
          chartData={chartData}
          data={data}
          search={search}
          setSearch={setSearch}
          statusFilt={statusFilt}
          toggleStatus={toggleStatus}
          clearStatus={clearStatus}
          contFilt={contFilt}
          toggleCont={toggleCont}
          clearCont={clearCont}
          boroughFilt={boroughFilt}
          toggleBorough={toggleBorough}
          clearBorough={clearBorough}
          contractors={contractors}
          boroughs={boroughs}
          filteredWOs={filteredWOs}
          visibleCount={visibleCount}
          onShowMore={() => setVisibleCount(c => c + WO_PAGE_SIZE)}
          onShowAll={() => setVisibleCount(filteredWOs.length)}
          completedTodayLabel={COMPLETED_TODAY}
          onDocsChange={onDocsChange}
          onChangeStatus={(wo) => setStatusPickerWO(wo)}
          onDeleteWO={(wo) => setDeleteWO(wo)}
        />
      )}

      {showDownloadModal && (
        <DownloadDocumentsModal
          contractors={contractors}
          onClose={() => setShowDownloadModal(false)}
        />
      )}

      {statusPickerWO && (
        <StatusPickerModal
          wo={statusPickerWO}
          onSaved={load}
          onClose={() => setStatusPickerWO(null)}
        />
      )}

      {deleteWO && (
        <DeleteWOModal
          wo={deleteWO}
          onDeleted={(result) => {
            // Optimistic: drop the row immediately so the dashboard
            // reflects the delete without waiting for the refetch.
            setData(prev => prev ? { ...prev, wos: prev.wos.filter(w => w.id !== deleteWO.id) } : prev)
            // Then refresh to pick up updated stats / counts from the server.
            load()
            // Eslint/no-undef: the toast helper isn't present in Dashboard; we
            // rely on the modal's own success state. Result counts (in
            // `result`) are surfaced via console for now.
            console.log('WO deleted:', result)
          }}
          onClose={() => setDeleteWO(null)}
        />
      )}
    </div>
  )
}

// ── Operations tab body ───────────────────────────────────────
// Extracted from Dashboard's render so the Revenue tab can swap in
// without re-rendering all the operations chrome. Pure presentation
// — every piece of state still lives in Dashboard().
function OperationsTabContent({
  stats, attention, chartData, data,
  search, setSearch,
  statusFilt, toggleStatus, clearStatus,
  contFilt, toggleCont, clearCont,
  boroughFilt, toggleBorough, clearBorough,
  contractors, boroughs,
  filteredWOs,
  visibleCount, onShowMore, onShowAll,
  completedTodayLabel,
  onChangeStatus, onDeleteWO,
  onDocsChange,
}) {
  return (
    <>
      {/* ── Stat cards ───────────────────────────────── */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <StatCard
          label="Total WOs"
          value={stats?.total}
          sub="all time"
          color="text-navy"
        />
        <StatCard
          label="In Progress"
          value={stats?.in_progress}
          sub={stats?.dispatched ? `+ ${stats.dispatched} dispatched` : undefined}
          color="text-orange-600"
        />
        <StatCard
          label="Needs Dispatch"
          value={stats?.received}
          sub="received, not dispatched"
          color="text-blue-600"
        />
        <StatCard
          label="Completed"
          value={stats?.complete}
          sub="all time"
          color="text-green-600"
        />
      </div>

      {/* ── Attention banner ─────────────────────────── */}
      {attention.length > 0 && (
        <div className="card p-4 border-amber-200 bg-amber-50">
          <div className="flex items-start gap-3">
            <span className="text-amber-500 text-xl flex-shrink-0">⚠</span>
            <div>
              <p className="text-sm font-bold text-amber-800">
                {attention.length} WO{attention.length !== 1 ? 's' : ''} need attention
              </p>
              <p className="text-xs text-amber-700 mt-0.5">
                {attention.slice(0, 5).join(', ')}
                {attention.length > 5 && ` +${attention.length - 5} more`}
              </p>
            </div>
          </div>
        </div>
      )}

      {/* ── Charts row ───────────────────────────────── */}
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        {/* Pipeline bar chart */}
        <div className="card p-4">
          <p className="section-label">Pipeline</p>
          <ResponsiveContainer width="100%" height={160}>
            <BarChart data={chartData} barSize={32}>
              <XAxis
                dataKey="name"
                tick={{ fontSize: 10, fill: '#94a3b8' }}
                axisLine={false}
                tickLine={false}
              />
              <YAxis
                allowDecimals={false}
                tick={{ fontSize: 10, fill: '#94a3b8' }}
                axisLine={false}
                tickLine={false}
                width={20}
              />
              <Tooltip
                contentStyle={{ fontSize: 12, borderRadius: 8, border: '1px solid #e2e8f0' }}
                cursor={{ fill: '#f1f5f9' }}
              />
              <Bar dataKey="value" radius={[4, 4, 0, 0]}>
                {chartData.map((entry, i) => (
                  <Cell key={i} fill={entry.fill} />
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* Contractor breakdown */}
        <div className="card p-4">
          <p className="section-label">By Contractor</p>
          <div className="space-y-2 mt-1">
            {Object.entries(data?.byContractor ?? {}).map(([name, count]) => {
              const pct = stats?.total ? Math.round((count / stats.total) * 100) : 0
              return (
                <div key={name}>
                  <div className="flex justify-between text-xs mb-1">
                    <span className="text-slate-700 font-medium truncate max-w-[70%]">{name}</span>
                    <span className="text-slate-500">{count} WOs</span>
                  </div>
                  <div className="h-1.5 bg-slate-100 rounded-full overflow-hidden">
                    <div
                      className="h-full bg-navy rounded-full transition-all duration-700"
                      style={{ width: pct + '%' }}
                    />
                  </div>
                </div>
              )
            })}
            {Object.keys(data?.byContractor ?? {}).length === 0 && (
              <p className="text-slate-400 text-sm">No data yet</p>
            )}
          </div>
        </div>
      </div>

      {/* ── WO Table ─────────────────────────────────── */}
      <div className="card overflow-hidden">
        <div className="p-4 border-b border-slate-100 space-y-3">
          <p className="section-label">Work Orders</p>

          {/* Search */}
          <input
            type="text"
            placeholder="Search WO #, location, or contractor…"
            value={search}
            onChange={e => setSearch(e.target.value)}
            className="field-input max-w-sm text-sm"
          />

          {/* Filters */}
          <div className="space-y-2">
            <FilterBar
              label="Status"
              options={['Received', 'Dispatched', 'In Progress', 'Completed', 'Returned', completedTodayLabel]}
              selected={statusFilt}
              onToggle={toggleStatus}
              onClear={clearStatus}
            />
            <FilterBar
              label="Contractor"
              options={contractors}
              selected={contFilt}
              onToggle={toggleCont}
              onClear={clearCont}
            />
            <FilterBar
              label="Borough"
              options={boroughs}
              selected={boroughFilt}
              onToggle={toggleBorough}
              onClear={clearBorough}
            />
          </div>

          <p className="text-xs text-slate-400">
            Showing {Math.min(visibleCount, filteredWOs.length)} of {filteredWOs.length} work orders
            {filteredWOs.length !== (data?.wos?.length ?? 0) && (
              <> · {data?.wos?.length ?? 0} total</>
            )}
          </p>
        </div>

        {/* Table — horizontally scrollable on mobile */}
        <div className="overflow-x-auto">
          <table className="w-full min-w-[820px]">
            <thead>
              <tr className="border-b border-slate-100 bg-slate-50">
                {['WO #', 'Contractor', 'Boro', 'Location', 'Due Date', 'Status', 'Quantity', 'Progress', 'Docs', 'Drive', ''].map((h, i) => (
                  <th
                    key={h || `col-${i}`}
                    className="py-2.5 px-3 text-left text-[10px] font-extrabold
                               uppercase tracking-wider text-slate-400"
                  >
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredWOs.length === 0 ? (
                <tr>
                  <td colSpan={11} className="py-12 text-center text-slate-400 text-sm">
                    No work orders match your filters.
                  </td>
                </tr>
              ) : (
                filteredWOs.slice(0, visibleCount).map(wo => (
                  <WORow
                    key={wo.id}
                    wo={wo}
                    flagged={attention.includes(wo.id)}
                    onDocsChange={onDocsChange}
                    onChangeStatus={onChangeStatus}
                    onDeleteWO={onDeleteWO}
                  />
                ))
              )}
            </tbody>
          </table>
        </div>

        {/* Expand footer — only render when there's more to reveal. */}
        {filteredWOs.length > visibleCount && (
          <div className="flex items-center justify-center gap-3 pt-3 text-sm">
            <button
              type="button"
              onClick={onShowMore}
              className="btn-outline text-xs px-4 py-1.5"
            >
              Show {Math.min(20, filteredWOs.length - visibleCount)} more
            </button>
            <button
              type="button"
              onClick={onShowAll}
              className="text-xs font-semibold text-navy hover:underline"
            >
              Show all {filteredWOs.length}
            </button>
          </div>
        )}
      </div>
    </>
  )
}
