import { useState, useEffect, useMemo } from 'react'
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
