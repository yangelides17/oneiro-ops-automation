import { useState, useEffect, useMemo } from 'react'
import { Link } from 'react-router-dom'
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer,
  CartesianGrid,
} from 'recharts'
import { opToday } from '../lib/dateOps'
import { displayCategory } from '../lib/markingCategories'

// ── Date range presets (mirror RevenueTab) ────────────────────
const PRESETS = [
  { id: '7d',     label: 'Last 7 days' },
  { id: '30d',    label: 'Last 30 days' },
  { id: 'mtd',    label: 'Month to date' },
  { id: 'last',   label: 'Last month' },
  { id: 'season', label: 'Season to date' },
  { id: 'custom', label: 'Custom' },
]

const isoOf = (d) => {
  const y = d.getFullYear()
  const m = String(d.getMonth() + 1).padStart(2, '0')
  const dd = String(d.getDate()).padStart(2, '0')
  return `${y}-${m}-${dd}`
}

// ── Weekly bucketing for long-range daily charts ──────────────
// Daily bars at >35 days start to crush — switch to weekly buckets
// (week-start = Sunday) summing the supplied numeric fields. Returns
// an array shaped like the input (same `date` field, summed values).
function bucketByWeek(daily, numericFields) {
  if (!daily || daily.length === 0) return []
  const buckets = {}
  const order = []
  daily.forEach(d => {
    const dt = new Date(d.date + 'T12:00:00')   // noon to dodge TZ DST edge
    const sunday = new Date(dt)
    sunday.setDate(dt.getDate() - dt.getDay())
    const key = isoOf(sunday)
    if (!buckets[key]) {
      const seed = { date: key }
      numericFields.forEach(f => { seed[f] = 0 })
      buckets[key] = seed
      order.push(key)
    }
    numericFields.forEach(f => {
      buckets[key][f] = (buckets[key][f] || 0) + (Number(d[f]) || 0)
    })
  })
  return order.map(k => buckets[k])
}

const WEEKLY_THRESHOLD = 35

function rangeForPreset(id, customStart, customEnd) {
  const today = new Date()
  if (id === '7d') {
    const start = new Date(today); start.setDate(today.getDate() - 6)
    return { start: isoOf(start), end: isoOf(today) }
  }
  if (id === '30d') {
    const start = new Date(today); start.setDate(today.getDate() - 29)
    return { start: isoOf(start), end: isoOf(today) }
  }
  if (id === 'mtd') {
    const start = new Date(today.getFullYear(), today.getMonth(), 1)
    return { start: isoOf(start), end: isoOf(today) }
  }
  if (id === 'last') {
    const lastMonthStart = new Date(today.getFullYear(), today.getMonth() - 1, 1)
    const lastMonthEnd   = new Date(today.getFullYear(), today.getMonth(),     0)
    return { start: isoOf(lastMonthStart), end: isoOf(lastMonthEnd) }
  }
  if (id === 'season') {
    // Striping season runs Jan 1 → today for this calendar year. Refine
    // later if Oneiro starts treating "season" as a different window.
    const start = new Date(today.getFullYear(), 0, 1)
    return { start: isoOf(start), end: isoOf(today) }
  }
  return { start: customStart || opToday(), end: customEnd || opToday() }
}

// ── Number / unit formatters ──────────────────────────────────
const fmtNum = (n) => (Number(n) || 0).toLocaleString('en-US', { maximumFractionDigits: 0 })
const fmtNumDecimal = (n) => (Number(n) || 0).toLocaleString('en-US', { maximumFractionDigits: 1 })

// Per-unit display metadata used across the tab. The user-facing
// title is the work category; the unit code is treated as a
// subheader.  Using a single map keeps tab labels in sync — change
// here and every chart / table picks it up.
const UNIT_META = {
  SF: { title: 'MMA Work',                 short: 'MMA Work' },
  LF: { title: 'Lines and Crosswalks',     short: 'Lines & Xwalks' },
  EA: { title: 'Msgs, Arrows and Symbols', short: 'Msgs / Arrows / Symbols' },
}

// ── Fetch ─────────────────────────────────────────────────────
async function fetchProduction(start, end) {
  const qs = new URLSearchParams()
  if (start) qs.set('start', start)
  if (end)   qs.set('end',   end)
  const res = await fetch('/api/production?' + qs.toString())
  if (!res.ok) throw new Error('HTTP ' + res.status)
  const data = await res.json()
  if (data.error) throw new Error(data.error)
  return data
}

// ── Range picker ──────────────────────────────────────────────
function RangePicker({ preset, onPresetChange, customStart, customEnd, onCustomChange }) {
  return (
    <div className="card p-4 space-y-3">
      <div className="flex flex-wrap items-center gap-2">
        <span className="text-xs font-semibold text-slate-500">Range</span>
        {PRESETS.map(p => (
          <button
            key={p.id}
            onClick={() => onPresetChange(p.id)}
            className={`text-xs px-3 py-1 rounded-full border font-medium transition-all
              ${preset === p.id
                ? 'bg-navy text-white border-navy'
                : 'bg-white text-slate-600 border-slate-200 hover:border-navy/40'
              }`}
          >
            {p.label}
          </button>
        ))}
      </div>
      {preset === 'custom' && (
        <div className="flex flex-wrap items-center gap-2 pt-2 border-t border-slate-100">
          <label className="text-xs font-semibold text-slate-500">Start</label>
          <input
            type="date"
            value={customStart}
            onChange={e => onCustomChange({ start: e.target.value, end: customEnd })}
            className="field-input"
          />
          <label className="text-xs font-semibold text-slate-500">End</label>
          <input
            type="date"
            value={customEnd}
            onChange={e => onCustomChange({ start: customStart, end: e.target.value })}
            className="field-input"
          />
        </div>
      )}
    </div>
  )
}

// ── KPI card ──────────────────────────────────────────────────
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

// ── Daily chart for a single unit ─────────────────────────────
function UnitDailyChart({ daily, unit, color }) {
  // Re-key the unit field to `qty` so bucketByWeek (which works off
  // field names) can sum it without needing a per-unit special case.
  const reKeyed = (daily || []).map(d => ({ date: d.date, qty: d[unit] || 0 }))
  const weekly  = reKeyed.length > WEEKLY_THRESHOLD
  const source  = weekly ? bucketByWeek(reKeyed, ['qty']) : reKeyed
  const data    = source.map(d => ({ date: d.date.slice(5), qty: d.qty }))
  const total = data.reduce((s, r) => s + r.qty, 0)
  const meta = UNIT_META[unit] || { title: unit, short: unit }
  const xInterval = data.length > 14 ? Math.max(0, Math.floor(data.length / 10) - 1) : 0
  return (
    <div className="card p-4">
      <div className="flex items-start justify-between mb-1 gap-2">
        <div>
          <p className="section-label leading-tight">{meta.title}</p>
          <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400 mt-0.5">
            {weekly ? 'Weekly' : 'Daily'} · {unit}
          </p>
        </div>
        <span className="text-xs text-slate-500 font-semibold whitespace-nowrap">
          {fmtNum(total)} {unit}
        </span>
      </div>
      <ResponsiveContainer width="100%" height={180}>
        <BarChart data={data} barSize={weekly ? 14 : (data.length > 31 ? 5 : 12)}>
          <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" vertical={false} />
          <XAxis dataKey="date" tick={{ fontSize: 10, fill: '#94a3b8' }} axisLine={false} tickLine={false} interval={xInterval} />
          <YAxis
            tick={{ fontSize: 10, fill: '#94a3b8' }}
            axisLine={false}
            tickLine={false}
            tickFormatter={fmtNum}
            width={48}
          />
          <Tooltip
            contentStyle={{ fontSize: 12, borderRadius: 8, border: '1px solid #e2e8f0' }}
            formatter={(v) => [fmtNum(v) + ' ' + unit, unit]}
            labelFormatter={(l) => weekly ? `Week of ${l}` : l}
            cursor={{ fill: '#f1f5f9' }}
          />
          <Bar dataKey="qty" radius={[3, 3, 0, 0]} fill={color} />
        </BarChart>
      </ResponsiveContainer>
    </div>
  )
}

// ── Contractor breakdown table ────────────────────────────────
// Two-line column headers — work category as the primary label, unit
// code as the subhead. Matches the tab-wide convention.
function UnitTh({ unit }) {
  const meta = UNIT_META[unit] || { short: unit }
  return (
    <th className="py-2.5 px-3 text-left text-[10px] font-extrabold
                   uppercase tracking-wider text-slate-500">
      <div className="leading-tight">{meta.short}</div>
      <div className="text-[9px] text-slate-400 font-bold mt-0.5">{unit}</div>
    </th>
  )
}

function ContractorTable({ rows }) {
  return (
    <div className="card overflow-hidden">
      <div className="p-4 border-b border-slate-100">
        <p className="section-label">Production by Contractor</p>
      </div>
      <div className="overflow-x-auto overscroll-x-contain">
        <table className="w-full min-w-[460px]">
          <thead>
            <tr className="border-b border-slate-100 bg-slate-50">
              <th className="py-2.5 px-3 text-left text-[10px] font-extrabold
                             uppercase tracking-wider text-slate-400">
                Contractor
              </th>
              <UnitTh unit="SF" />
              <UnitTh unit="LF" />
              <UnitTh unit="EA" />
              <th className="py-2.5 px-3 text-left text-[10px] font-extrabold
                             uppercase tracking-wider text-slate-400">
                Items
              </th>
            </tr>
          </thead>
          <tbody>
            {(!rows || rows.length === 0) ? (
              <tr>
                <td colSpan={5} className="py-8 text-center text-slate-400 text-sm">
                  No production in this range.
                </td>
              </tr>
            ) : rows.map(r => (
              <tr key={r.contractor} className="border-b border-slate-100 text-sm hover:bg-slate-50">
                <td className="py-2.5 px-3 text-slate-700 font-medium">{r.contractor}</td>
                <td className="py-2.5 px-3 text-slate-600 whitespace-nowrap">{fmtNum(r.SF)}</td>
                <td className="py-2.5 px-3 text-slate-600 whitespace-nowrap">{fmtNum(r.LF)}</td>
                <td className="py-2.5 px-3 text-slate-600 whitespace-nowrap">{fmtNum(r.EA)}</td>
                <td className="py-2.5 px-3 text-slate-500 text-xs">{r.items}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}

// ── By-category breakdown grouped by unit ─────────────────────
function CategoryBreakdown({ rows }) {
  const grouped = useMemo(() => {
    const out = { SF: [], LF: [], EA: [] }
    ;(rows || []).forEach(r => {
      if (out[r.unit]) out[r.unit].push(r)
    })
    return out
  }, [rows])

  const sections = ['SF', 'LF', 'EA']

  return (
    <div className="card p-4">
      <p className="section-label mb-3">By Marking Type</p>
      <div className="space-y-4">
        {sections.map(unit => {
          const list = grouped[unit] || []
          const max  = list[0]?.qty || 0
          const meta = UNIT_META[unit] || { title: unit }
          return (
            <div key={unit}>
              <div className="mb-2">
                <p className="text-xs font-semibold text-slate-700 leading-tight">
                  {meta.title}
                </p>
                <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400 mt-0.5">
                  {unit}
                </p>
              </div>
              {list.length === 0 ? (
                <p className="text-slate-400 text-sm">No {unit} markings in this range</p>
              ) : (
                <div className="space-y-1.5">
                  {list.slice(0, 12).map(r => {
                    const pct = max > 0 ? Math.round((r.qty / max) * 100) : 0
                    return (
                      <div key={r.category + '|' + r.unit}>
                        <div className="flex justify-between text-xs mb-0.5">
                          <span className="text-slate-700 font-medium truncate max-w-[60%]">
                            {displayCategory(r.category)}
                          </span>
                          <span className="text-slate-500 whitespace-nowrap">
                            {fmtNum(r.qty)} {r.unit} · {r.items} item{r.items === 1 ? '' : 's'}
                          </span>
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
                </div>
              )}
            </div>
          )
        })}
      </div>
    </div>
  )
}

// ── Top WOs table ─────────────────────────────────────────────
function TopWosTable({ topWos }) {
  if (!topWos || topWos.length === 0) {
    return (
      <div className="card p-4">
        <p className="section-label">Top WOs</p>
        <p className="text-slate-400 text-sm mt-2">No production in this range</p>
      </div>
    )
  }
  return (
    <div className="card overflow-hidden">
      <div className="p-4 border-b border-slate-100">
        <p className="section-label">Top WOs by Production</p>
      </div>
      <div className="overflow-x-auto overscroll-x-contain">
        <table className="w-full min-w-[640px]">
          <thead>
            <tr className="border-b border-slate-100 bg-slate-50">
              {['WO #', 'Contractor', 'Location'].map(h => (
                <th key={h} className="py-2.5 px-3 text-left text-[10px] font-extrabold
                                       uppercase tracking-wider text-slate-400">
                  {h}
                </th>
              ))}
              <UnitTh unit="SF" />
              <UnitTh unit="LF" />
              <UnitTh unit="EA" />
              <th className="py-2.5 px-3 text-left text-[10px] font-extrabold
                             uppercase tracking-wider text-slate-400">
                Items
              </th>
            </tr>
          </thead>
          <tbody>
            {topWos.map(w => (
              <tr key={w.wo_id} className="border-b border-slate-100 text-sm hover:bg-slate-50">
                <td className="py-2.5 px-3 font-mono font-semibold whitespace-nowrap">
                  <Link to={`/field-report?wo=${encodeURIComponent(w.wo_id)}`} className="text-navy hover:underline">
                    {w.wo_id}
                  </Link>
                </td>
                <td className="py-2.5 px-3 text-slate-700">{w.contractor || '—'}</td>
                <td className="py-2.5 px-3 text-slate-600 max-w-[180px] truncate">{w.location || '—'}</td>
                <td className="py-2.5 px-3 text-slate-600 whitespace-nowrap">{fmtNum(w.SF)}</td>
                <td className="py-2.5 px-3 text-slate-600 whitespace-nowrap">{fmtNum(w.LF)}</td>
                <td className="py-2.5 px-3 text-slate-600 whitespace-nowrap">{fmtNum(w.EA)}</td>
                <td className="py-2.5 px-3 text-slate-500 text-xs">{w.items}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}

// ── Main ──────────────────────────────────────────────────────
export default function ProductionTab() {
  const [preset, setPreset]           = useState('season')
  const [custom, setCustom]           = useState({ start: '', end: '' })
  const [data,    setData]            = useState(null)
  const [loading, setLoading]         = useState(true)
  const [error,   setError]           = useState(null)
  const [lastRefresh, setLastRefresh] = useState(null)

  const range = useMemo(
    () => rangeForPreset(preset, custom.start, custom.end),
    [preset, custom.start, custom.end]
  )

  const load = async () => {
    setLoading(true)
    setError(null)
    try {
      const d = await fetchProduction(range.start, range.end)
      setData(d)
      setLastRefresh(new Date())
    } catch (e) {
      setError(e.message)
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => { load() }, [range.start, range.end])

  if (loading && !data) {
    return (
      <div className="flex items-center justify-center min-h-[40vh]">
        <div className="flex flex-col items-center gap-4">
          <div className="w-10 h-10 border-[3px] border-slate-200 border-t-navy
                          rounded-full animate-spin" />
          <p className="text-slate-400 text-sm">Loading production…</p>
        </div>
      </div>
    )
  }

  if (error && !data) {
    return (
      <div className="flex items-center justify-center min-h-[40vh]">
        <div className="text-center max-w-sm">
          <div className="text-4xl mb-4">⚠️</div>
          <h2 className="text-lg font-bold text-slate-800 mb-2">Failed to load production</h2>
          <p className="text-slate-500 text-sm mb-4">{error}</p>
          <button onClick={load} className="btn-outline text-sm px-4 py-2">Try again</button>
        </div>
      </div>
    )
  }

  const totals       = data?.totals       || { SF: 0, LF: 0, EA: 0, items: 0 }
  const shifts       = data?.shifts       || { count: 0, days_in_range: 0, pct_days_worked: 0, longest_streak: 0 }
  const daily        = data?.daily        || []
  const byContractor = data?.by_contractor|| []
  const byCategory   = data?.by_category  || []
  const topWos       = data?.top_wos      || []

  return (
    <div className="space-y-6">
      <RangePicker
        preset={preset}
        onPresetChange={setPreset}
        customStart={custom.start}
        customEnd={custom.end}
        onCustomChange={setCustom}
      />

      {lastRefresh && (
        <p className="text-slate-400 text-xs -mt-3">
          Range: {range.start} → {range.end} · Updated {lastRefresh.toLocaleTimeString()}
        </p>
      )}

      {/* Quantity KPIs — title is the work category, unit (SF/LF/EA) is the subhead */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <StatCard label={UNIT_META.SF.title} value={fmtNum(totals.SF)} sub="SF" color="text-navy" />
        <StatCard label={UNIT_META.LF.title} value={fmtNum(totals.LF)} sub="LF" color="text-blue-600" />
        <StatCard label={UNIT_META.EA.title} value={fmtNum(totals.EA)} sub="EA" color="text-orange-600" />
        <StatCard label="Items"              value={fmtNum(totals.items)} sub="completed" color="text-green-600" />
      </div>

      {/* Shift KPIs */}
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
        <StatCard
          label="Shifts Worked"
          value={fmtNum(shifts.count)}
          sub={`of ${fmtNum(shifts.days_in_range)} days in range`}
          color="text-navy"
        />
        <StatCard
          label="% Days Worked"
          value={fmtNumDecimal(shifts.pct_days_worked) + '%'}
          sub={shifts.pct_days_worked >= 50 ? 'half the days, easy' : 'in this range'}
          color="text-blue-600"
        />
        <StatCard
          label="Longest Streak"
          value={fmtNum(shifts.longest_streak) + (shifts.longest_streak === 1 ? ' day' : ' days')}
          sub="consecutive shifts"
          color={shifts.longest_streak >= 5 ? 'text-green-600' : 'text-slate-700'}
        />
      </div>

      {/* Daily charts — 3-up on desktop, stacked on mobile */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <UnitDailyChart daily={daily} unit="SF" color="#1e40af" />
        <UnitDailyChart daily={daily} unit="LF" color="#2563eb" />
        <UnitDailyChart daily={daily} unit="EA" color="#f97316" />
      </div>

      {/* Breakdown grid */}
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <ContractorTable rows={byContractor} />
        <CategoryBreakdown rows={byCategory} />
      </div>

      <TopWosTable topWos={topWos} />
    </div>
  )
}
