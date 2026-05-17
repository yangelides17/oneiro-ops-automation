import { useState, useEffect, useMemo } from 'react'
import { Link } from 'react-router-dom'
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell,
  CartesianGrid, Legend,
} from 'recharts'
import { opToday } from '../lib/dateOps'
import {
  PRICING_GROUPS,
  PRICING_GROUP_LABEL,
  NEEDS_PRICING_REASON_LABEL,
} from '../lib/pricing'

// ── Date range helpers ────────────────────────────────────────
// Range picker presets: server defaults to MTD when start/end are
// blank, but we explicitly send dates so the UI label matches what
// the server saw.
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
    // Striping season runs Jan 1 → today for this calendar year. Mirror
    // of the ProductionTab "season" preset — keep both tabs in sync.
    const start = new Date(today.getFullYear(), 0, 1)
    return { start: isoOf(start), end: isoOf(today) }
  }
  return { start: customStart || opToday(), end: customEnd || opToday() }
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

const WEEKLY_THRESHOLD = 35   // shared across both charts in this tab

// ── Currency / percent formatters ─────────────────────────────
const fmtUsd = (n) =>
  (Number(n) || 0).toLocaleString('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 })
const fmtUsdSmall = (n) =>
  (Number(n) || 0).toLocaleString('en-US', { style: 'currency', currency: 'USD' })

// Group fill colors — mirror the existing dashboard palette so the
// Revenue tab feels like a continuation of the Operations tab.
const GROUP_COLOR = {
  line4:         '#1e40af',  // navy-ish
  line12:        '#2563eb',  // blue
  preformed:     '#7c3aed',  // violet
  extruded:      '#f97316',  // orange
  color_surface: '#16a34a',  // green
}

// ── Fetch ─────────────────────────────────────────────────────
async function fetchRevenue(start, end) {
  const qs = new URLSearchParams()
  if (start) qs.set('start', start)
  if (end)   qs.set('end',   end)
  const res = await fetch('/api/revenue?' + qs.toString())
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
            className="field-input text-sm"
          />
          <label className="text-xs font-semibold text-slate-500">End</label>
          <input
            type="date"
            value={customEnd}
            onChange={e => onCustomChange({ start: customStart, end: e.target.value })}
            className="field-input text-sm"
          />
        </div>
      )}
    </div>
  )
}

// ── KPI cards ─────────────────────────────────────────────────
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

// ── Charts ────────────────────────────────────────────────────
function DailyRevenueChart({ daily }) {
  const weekly = (daily?.length || 0) > WEEKLY_THRESHOLD
  const source = weekly ? bucketByWeek(daily, ['revenue']) : (daily || [])
  const data = source.map(d => ({
    date: d.date.slice(5),  // MM-DD only on the X axis
    revenue: d.revenue,
  }))
  // Cap label density so a ~20-bar season chart doesn't shout 20 dates
  // at the reader; recharts hides overflows via auto-thin, but forcing
  // an interval guarantees readable spacing.
  const xInterval = data.length > 14 ? Math.max(0, Math.floor(data.length / 12) - 1) : 0
  return (
    <div className="card p-4">
      <div className="flex items-baseline justify-between">
        <p className="section-label">{weekly ? 'Weekly Revenue' : 'Daily Revenue'}</p>
        {weekly && (
          <span className="text-[10px] uppercase tracking-widest text-slate-400 font-bold">
            bucketed by week
          </span>
        )}
      </div>
      <ResponsiveContainer width="100%" height={220}>
        <BarChart data={data} barSize={weekly ? 18 : (data.length > 31 ? 6 : 16)}>
          <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" vertical={false} />
          <XAxis dataKey="date" tick={{ fontSize: 10, fill: '#94a3b8' }} axisLine={false} tickLine={false} interval={xInterval} />
          <YAxis
            tick={{ fontSize: 10, fill: '#94a3b8' }}
            axisLine={false}
            tickLine={false}
            tickFormatter={fmtUsd}
            width={56}
          />
          <Tooltip
            contentStyle={{ fontSize: 12, borderRadius: 8, border: '1px solid #e2e8f0' }}
            formatter={(v) => fmtUsdSmall(v)}
            labelFormatter={(l) => weekly ? `Week of ${l}` : l}
            cursor={{ fill: '#f1f5f9' }}
          />
          <Bar dataKey="revenue" radius={[4, 4, 0, 0]} fill="#1e40af" />
        </BarChart>
      </ResponsiveContainer>
    </div>
  )
}

function GroupStackedChart({ daily }) {
  const weekly = (daily?.length || 0) > WEEKLY_THRESHOLD
  // Recharts likes flat objects per X tick — flatten by_group out.
  // For the weekly view we flatten first, then sum per pricing group.
  const flat = (daily || []).map(d => {
    const row = { date: d.date }
    PRICING_GROUPS.forEach(g => { row[g] = (d.by_group && d.by_group[g]) || 0 })
    return row
  })
  const source = weekly ? bucketByWeek(flat, PRICING_GROUPS) : flat
  const data = source.map(d => ({ ...d, date: d.date.slice(5) }))
  const xInterval = data.length > 14 ? Math.max(0, Math.floor(data.length / 12) - 1) : 0
  return (
    <div className="card p-4">
      <div className="flex items-baseline justify-between">
        <p className="section-label">
          {weekly ? 'Weekly Revenue by Pricing Group' : 'Daily Revenue by Pricing Group'}
        </p>
        {weekly && (
          <span className="text-[10px] uppercase tracking-widest text-slate-400 font-bold">
            bucketed by week
          </span>
        )}
      </div>
      <ResponsiveContainer width="100%" height={240}>
        <BarChart data={data} barSize={weekly ? 18 : (data.length > 31 ? 6 : 16)}>
          <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" vertical={false} />
          <XAxis dataKey="date" tick={{ fontSize: 10, fill: '#94a3b8' }} axisLine={false} tickLine={false} interval={xInterval} />
          <YAxis
            tick={{ fontSize: 10, fill: '#94a3b8' }}
            axisLine={false}
            tickLine={false}
            tickFormatter={fmtUsd}
            width={56}
          />
          <Tooltip
            contentStyle={{ fontSize: 12, borderRadius: 8, border: '1px solid #e2e8f0' }}
            formatter={(v, name) => [fmtUsdSmall(v), PRICING_GROUP_LABEL[name] || name]}
            labelFormatter={(l) => weekly ? `Week of ${l}` : l}
            cursor={{ fill: '#f1f5f9' }}
          />
          <Legend
            wrapperStyle={{ fontSize: 11 }}
            formatter={(value) => PRICING_GROUP_LABEL[value] || value}
          />
          {PRICING_GROUPS.map(g => (
            <Bar
              key={g}
              dataKey={g}
              stackId="rev"
              fill={GROUP_COLOR[g] || '#64748b'}
            />
          ))}
        </BarChart>
      </ResponsiveContainer>
    </div>
  )
}

function ContractorBreakdown({ byContractor, total }) {
  const list = byContractor || []
  return (
    <div className="card p-4">
      <p className="section-label">Revenue by Contractor</p>
      <div className="space-y-2 mt-1">
        {list.length === 0 && <p className="text-slate-400 text-sm">No data in range</p>}
        {list.map(row => {
          const pct = total > 0 ? Math.round((row.revenue / total) * 100) : 0
          return (
            <div key={row.contractor}>
              <div className="flex justify-between text-xs mb-1">
                <span className="text-slate-700 font-medium truncate max-w-[60%]">
                  {row.contractor}
                </span>
                <span className="text-slate-500">{fmtUsdSmall(row.revenue)}</span>
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
    </div>
  )
}

function TopWosTable({ topWos }) {
  if (!topWos || topWos.length === 0) {
    return (
      <div className="card p-4">
        <p className="section-label">Top WOs</p>
        <p className="text-slate-400 text-sm mt-2">No revenue in this range</p>
      </div>
    )
  }
  return (
    <div className="card overflow-hidden">
      <div className="p-4 border-b border-slate-100">
        <p className="section-label">Top WOs by Revenue</p>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full min-w-[520px]">
          <thead>
            <tr className="border-b border-slate-100 bg-slate-50">
              {['WO #', 'Contractor', 'Location', 'Items', 'Revenue'].map(h => (
                <th key={h} className="py-2.5 px-3 text-left text-[10px] font-extrabold
                                       uppercase tracking-wider text-slate-400">
                  {h}
                </th>
              ))}
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
                <td className="py-2.5 px-3 text-slate-600 max-w-[220px] truncate">{w.location || '—'}</td>
                <td className="py-2.5 px-3 text-slate-500 text-xs">{w.items}</td>
                <td className="py-2.5 px-3 text-slate-800 font-semibold whitespace-nowrap">
                  {fmtUsdSmall(w.revenue)}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}

function NeedsPricingPanel({ items }) {
  const [open, setOpen] = useState(true)
  if (!items || items.length === 0) return null

  // Group by reason for clarity.
  const byReason = {}
  items.forEach(it => {
    const k = it.reason || 'unknown'
    if (!byReason[k]) byReason[k] = []
    byReason[k].push(it)
  })

  return (
    <div className="card border-amber-200 bg-amber-50">
      <button
        onClick={() => setOpen(o => !o)}
        className="w-full p-4 flex items-start gap-3 text-left"
      >
        <span className="text-amber-500 text-xl flex-shrink-0">⚠</span>
        <div className="flex-1">
          <p className="text-sm font-bold text-amber-800">
            {items.length} item{items.length === 1 ? '' : 's'} need pricing review
          </p>
          <p className="text-xs text-amber-700 mt-0.5">
            Click to {open ? 'hide' : 'show'} details
          </p>
        </div>
        <span className="text-amber-600">{open ? '▴' : '▾'}</span>
      </button>
      {open && (
        <div className="px-4 pb-4 space-y-4">
          {Object.keys(byReason).map(reason => (
            <div key={reason}>
              <p className="text-xs font-semibold text-amber-900 mb-1.5">
                {NEEDS_PRICING_REASON_LABEL[reason] || reason} ({byReason[reason].length})
              </p>
              <div className="overflow-x-auto bg-white rounded-lg border border-amber-200">
                <table className="w-full min-w-[480px] text-xs">
                  <thead>
                    <tr className="border-b border-slate-100 bg-slate-50">
                      {['WO #', 'Item ID', 'Category', 'Qty', 'Unit'].map(h => (
                        <th key={h} className="py-2 px-3 text-left font-extrabold
                                               uppercase tracking-wider text-slate-400">
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {byReason[reason].map(it => (
                      <tr key={it.item_id} className="border-b border-slate-100 last:border-0">
                        <td className="py-1.5 px-3 font-mono">
                          <Link
                            to={`/field-report?wo=${encodeURIComponent(it.wo_id)}`}
                            className="text-navy hover:underline"
                          >
                            {it.wo_id}
                          </Link>
                        </td>
                        <td className="py-1.5 px-3 font-mono text-slate-500">{it.item_id}</td>
                        <td className="py-1.5 px-3 text-slate-700">{it.category}</td>
                        <td className="py-1.5 px-3 text-slate-600">{it.qty}</td>
                        <td className="py-1.5 px-3 text-slate-500">{it.unit}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

// ── Main ──────────────────────────────────────────────────────
export default function RevenueTab() {
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
      const d = await fetchRevenue(range.start, range.end)
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
          <p className="text-slate-400 text-sm">Loading revenue…</p>
        </div>
      </div>
    )
  }

  if (error && !data) {
    return (
      <div className="flex items-center justify-center min-h-[40vh]">
        <div className="text-center max-w-sm">
          <div className="text-4xl mb-4">⚠️</div>
          <h2 className="text-lg font-bold text-slate-800 mb-2">Failed to load revenue</h2>
          <p className="text-slate-500 text-sm mb-4">{error}</p>
          <button onClick={load} className="btn-outline text-sm px-4 py-2">Try again</button>
        </div>
      </div>
    )
  }

  const totals       = data?.totals       || { revenue: 0, items: 0, unpriced_items: 0 }
  const daily        = data?.daily        || []
  const byContractor = data?.by_contractor|| []
  const topWos       = data?.top_wos      || []
  const needs        = data?.needs_pricing|| []

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

      <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
        <StatCard
          label="Revenue"
          value={fmtUsd(totals.revenue)}
          sub={`${totals.items} priced item${totals.items === 1 ? '' : 's'}`}
          color="text-navy"
        />
        <StatCard
          label="Items Completed"
          value={totals.items}
          sub="in range"
          color="text-blue-600"
        />
        <StatCard
          label="Needs Pricing"
          value={totals.unpriced_items}
          sub={totals.unpriced_items === 0 ? 'all clear' : 'see panel below'}
          color={totals.unpriced_items === 0 ? 'text-green-600' : 'text-amber-600'}
        />
      </div>

      <NeedsPricingPanel items={needs} />

      <DailyRevenueChart daily={daily} />

      <GroupStackedChart daily={daily} />

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <ContractorBreakdown byContractor={byContractor} total={totals.revenue} />
        <TopWosTable topWos={topWos} />
      </div>
    </div>
  )
}
