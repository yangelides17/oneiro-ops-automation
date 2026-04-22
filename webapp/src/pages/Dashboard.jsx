import { useState, useEffect, useMemo, useRef } from 'react'
import { Link } from 'react-router-dom'
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell
} from 'recharts'
import StatusBadge from '../components/StatusBadge'

// ── API ───────────────────────────────────────────────────────
async function fetchDashboard() {
  const res = await fetch('/api/dashboard')
  if (!res.ok) throw new Error('HTTP ' + res.status)
  const data = await res.json()
  if (data.error) throw new Error(data.error)
  return data
}

// ── Date helpers ──────────────────────────────────────────────
const isoToday = () => {
  const d = new Date()
  return [d.getFullYear(), String(d.getMonth()+1).padStart(2,'0'), String(d.getDate()).padStart(2,'0')].join('-')
}
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
  const [busy,     setBusy]     = useState(false)
  const [toast,    setToast]    = useState(null)   // { ok, msg }
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

  // Auto-dismiss the toast after 6s so it doesn't pile up if the
  // user clicks several items in succession.
  useEffect(() => {
    if (!toast) return
    const t = setTimeout(() => setToast(null), 6000)
    return () => clearTimeout(t)
  }, [toast])

  async function runDailyDocs(isoDate) {
    setBusy(true); setOpen(false); setPicker(null); setToast(null)
    try {
      const res = await fetch('/api/tools/daily-documents', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ date: isoDate || '' }),
      })
      const data = await res.json().catch(() => ({}))
      if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
      const when = isoDate ? prettyDate(isoDate) : 'today'
      const n    = data.entries_found ?? 0
      setToast({
        ok: true,
        msg: n > 0
          ? `Daily documents generated for ${when} — ${n} sign-in entr${n === 1 ? 'y' : 'ies'} processed. Check "Docs Needing Review" in Drive.`
          : `No sign-in entries found for ${when}. Nothing generated.`,
      })
    } catch (err) {
      setToast({ ok: false, msg: err.message || 'Failed to generate daily documents' })
    } finally {
      setBusy(false)
    }
  }

  async function runProcessApproved() {
    setBusy(true); setOpen(false); setPicker(null); setToast(null)
    try {
      const res = await fetch('/api/tools/process-approved-docs', { method: 'POST' })
      const data = await res.json().catch(() => ({}))
      if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
      const archived = data.archived ?? 0
      const errored  = data.errored ?? 0
      let msg
      if (data.skipped) {
        msg = 'Another run is already in progress — it will finish shortly. No action taken.'
      } else if (archived === 0 && errored === 0) {
        msg = 'No pending docs in Approved Docs folder — nothing to process.'
      } else {
        const parts = []
        if (archived > 0) parts.push(`archived ${archived}`)
        if (errored > 0)  parts.push(`${errored} failed → Archive Errors`)
        msg = `Processed approved docs: ${parts.join(', ')}.`
      }
      setToast({ ok: errored === 0, msg })
    } catch (err) {
      setToast({ ok: false, msg: err.message || 'Failed to process approved docs' })
    } finally {
      setBusy(false)
    }
  }

  async function runCertPayroll(isoWeekStart) {
    setBusy(true); setOpen(false); setPicker(null); setToast(null)
    try {
      const res = await fetch('/api/tools/certified-payroll', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ week_start: isoWeekStart }),
      })
      const data = await res.json().catch(() => ({}))
      if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
      const n = data.contract_groups ?? 0
      setToast({
        ok: true,
        msg: n > 0
          ? `Certified payroll generated for week of ${prettyDate(isoWeekStart)} — ${n} contract group${n === 1 ? '' : 's'}. Check "Docs Needing Review → Certified Payroll" in Drive.`
          : `No sign-in entries found for week of ${prettyDate(isoWeekStart)}. Nothing generated.`,
      })
    } catch (err) {
      setToast({ ok: false, msg: err.message || 'Failed to generate certified payroll' })
    } finally {
      setBusy(false)
    }
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
        disabled={busy}
        onClick={() => setOpen(v => !v)}
        className="btn-ghost flex items-center gap-1.5 disabled:opacity-50"
      >
        {busy ? <span className="animate-spin inline-block">⟳</span> : <span>🛠️</span>}
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
                disabled={!pickerVal || busy}
                onClick={submitPicker}
                className="flex-1 py-2.5 rounded-xl text-sm font-bold bg-navy text-white hover:opacity-90 disabled:opacity-50"
              >
                Generate
              </button>
            </div>
          </div>
        </div>
      )}

      {toast && (
        <div className={`fixed bottom-6 right-6 z-50 max-w-md px-4 py-3 rounded-xl shadow-lg text-sm leading-snug
          ${toast.ok
            ? 'bg-green-50 border border-green-200 text-green-800'
            : 'bg-red-50 border border-red-200 text-red-800'}`}>
          <div className="flex items-start justify-between gap-3">
            <span>{toast.msg}</span>
            <button
              type="button"
              onClick={() => setToast(null)}
              className="opacity-60 hover:opacity-100"
              aria-label="Dismiss"
            >✕</button>
          </div>
        </div>
      )}
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
const ALL = 'All'

function FilterBar({ label, options, value, onChange }) {
  return (
    <div className="flex items-center gap-2 flex-wrap">
      <span className="text-xs font-semibold text-slate-500">{label}</span>
      {[ALL, ...options].map(opt => (
        <button
          key={opt}
          onClick={() => onChange(opt)}
          className={`text-xs px-3 py-1 rounded-full border font-medium transition-all
            ${value === opt
              ? 'bg-navy text-white border-navy'
              : 'bg-white text-slate-600 border-slate-200 hover:border-navy/40'
            }`}
        >
          {opt}
        </button>
      ))}
    </div>
  )
}

// ── WO Table row ──────────────────────────────────────────────
function WORow({ wo, flagged }) {
  const isOverdue = wo.due_date && new Date(wo.due_date) < new Date()
    && wo.status.toLowerCase() !== 'completed'

  return (
    <tr className={`border-b border-slate-100 text-sm transition-colors
                    hover:bg-slate-50 ${flagged ? 'bg-amber-50/60' : ''}`}>
      <td className="py-2.5 px-3 font-mono font-semibold text-navy whitespace-nowrap">
        {wo.id}
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
        {wo.due_date || '—'}
        {isOverdue && ' ⚡'}
      </td>
      <td className="py-2.5 px-3"><StatusBadge status={wo.status} /></td>
      <td className="py-2.5 px-3 text-slate-400 text-xs">{wo.sqft || '—'}</td>
      <td className="py-2.5 px-3">
        {wo.photos === 'Yes'
          ? <span className="text-green-600 text-xs font-semibold">✓ Yes</span>
          : <span className="text-slate-300 text-xs">—</span>}
      </td>
    </tr>
  )
}

// ── Main ──────────────────────────────────────────────────────
export default function Dashboard() {
  const [data,       setData]       = useState(null)
  const [loading,    setLoading]    = useState(true)
  const [error,      setError]      = useState(null)
  const [statusFilt, setStatusFilt] = useState(ALL)
  const [contFilt,   setContFilt]   = useState(ALL)
  const [boroughFilt,setBoroughFilt]= useState(ALL)
  const [search,     setSearch]     = useState('')
  const [lastRefresh,setLastRefresh]= useState(null)

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

  // Filter + search WOs
  const filteredWOs = useMemo(() => {
    if (!data?.wos) return []
    return data.wos.filter(wo => {
      if (statusFilt  !== ALL && wo.status    !== statusFilt)  return false
      if (contFilt    !== ALL && wo.contractor !== contFilt)   return false
      if (boroughFilt !== ALL && wo.borough    !== boroughFilt) return false
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
          <h1 className="text-2xl font-black text-navy leading-none">Operations Dashboard</h1>
          {lastRefresh && (
            <p className="text-slate-400 text-xs mt-1">
              Updated {lastRefresh.toLocaleTimeString()}
            </p>
          )}
        </div>
        <div className="flex items-center gap-2">
          <ToolsMenu />
          <button
            onClick={load}
            disabled={loading}
            className="btn-ghost flex items-center gap-1.5"
          >
            <span className={loading ? 'animate-spin inline-block' : ''}>↻</span>
            Refresh
          </button>
          <Link to="/field-report" className="btn-primary text-sm px-4 py-2">
            + Field Report
          </Link>
        </div>
      </div>

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
              options={['Received', 'Dispatched', 'In Progress', 'Completed']}
              value={statusFilt}
              onChange={setStatusFilt}
            />
            <FilterBar
              label="Contractor"
              options={contractors}
              value={contFilt}
              onChange={setContFilt}
            />
            <FilterBar
              label="Borough"
              options={boroughs}
              value={boroughFilt}
              onChange={setBoroughFilt}
            />
          </div>

          <p className="text-xs text-slate-400">
            Showing {filteredWOs.length} of {data?.wos?.length ?? 0} work orders
          </p>
        </div>

        {/* Table — horizontally scrollable on mobile */}
        <div className="overflow-x-auto">
          <table className="w-full min-w-[700px]">
            <thead>
              <tr className="border-b border-slate-100 bg-slate-50">
                {['WO #', 'Contractor', 'Boro', 'Location', 'Due Date', 'Status', 'SQFT', 'Photos'].map(h => (
                  <th
                    key={h}
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
                  <td colSpan={8} className="py-12 text-center text-slate-400 text-sm">
                    No work orders match your filters.
                  </td>
                </tr>
              ) : (
                filteredWOs.map(wo => (
                  <WORow
                    key={wo.id}
                    wo={wo}
                    flagged={attention.includes(wo.id)}
                  />
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

    </div>
  )
}
