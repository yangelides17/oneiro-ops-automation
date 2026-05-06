import { useEffect, useMemo, useRef, useState } from 'react'

/**
 * DocStatusTab — calendar view of Production Log + Sign-In (per day)
 * and Certified Payroll (per week). Uses the Doc Lifecycle Log as
 * the source of truth for per-doc state.
 *
 * UI layout:
 *   - Month nav strip (prev / month label / next / Today)
 *   - Two-column grid (stacks on mobile):
 *       Left  → Day calendar (PL + SI rolled up per cell)
 *       Right → Week calendar (CP per cell)
 *   - Pending lists below each calendar (FIFO oldest-first, all-time)
 *
 * Cell colors:
 *   gray   — work happened, nothing done
 *   amber  — partial (some doc_done / doc_sent missing)
 *   green  — fully done + sent
 *
 * Hover: ring on clickable cells.
 * Click: popover with per-(contract, borough) breakdown + Done/Sent
 * toggles. Optimistic update + revert on failure.
 */

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
]

const todayIso = () => {
  const d = new Date()
  const yy = d.getFullYear()
  const mm = String(d.getMonth() + 1).padStart(2, '0')
  const dd = String(d.getDate()).padStart(2, '0')
  return `${yy}-${mm}-${dd}`
}
const todayMonthIso = () => todayIso().slice(0, 7)

const fmtMonth = (monthIso) => {
  const [y, m] = monthIso.split('-').map(Number)
  return `${MONTHS[m - 1]} ${y}`
}
const shiftMonth = (monthIso, delta) => {
  const [y, m] = monthIso.split('-').map(Number)
  const d = new Date(y, m - 1 + delta, 1)
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`
}

const fmtShortDate = (iso) => {
  const [y, m, d] = String(iso || '').split('-').map(Number)
  if (!y || !m || !d) return iso
  return `${MONTHS[m - 1].slice(0, 3)} ${d}, ${y}`
}

// ── Cell background per status ────────────────────────────────
const STATUS_BG = {
  gray:  'bg-slate-100  text-slate-500 border-slate-200',
  amber: 'bg-amber-100  text-amber-800 border-amber-200',
  green: 'bg-green-100  text-green-800 border-green-200',
}

// ── Fetch ─────────────────────────────────────────────────────
async function fetchDocStatus(monthIso) {
  const res = await fetch('/api/doc-status?month=' + encodeURIComponent(monthIso))
  if (!res.ok) throw new Error('HTTP ' + res.status)
  const data = await res.json()
  if (data.error) throw new Error(data.error)
  return data
}

async function flipDocFlags(updates) {
  const res = await fetch('/api/doc-status/flags', {
    method:  'POST',
    headers: { 'Content-Type': 'application/json' },
    body:    JSON.stringify({ updates }),
  })
  const body = await res.json().catch(() => ({}))
  if (!res.ok || body.error) throw new Error(body.error || `HTTP ${res.status}`)
  return body
}

// ── Toggle pill (matches DocStatusChips's ToggleButton flavor) ─
function TogglePill({ label, on, disabled, pending, onClick, color }) {
  const base = 'text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded border transition-all'
  let style
  if (disabled) {
    style = 'bg-slate-50 text-slate-300 border-slate-100 cursor-not-allowed'
  } else if (on) {
    style = color === 'green'
      ? 'bg-green-500 text-white border-green-500 hover:bg-green-600'
      : 'bg-amber-500 text-white border-amber-500 hover:bg-amber-600'
  } else {
    style = 'bg-white text-slate-500 border-slate-200 hover:border-slate-400'
  }
  return (
    <button
      type="button"
      onClick={disabled ? undefined : onClick}
      disabled={disabled || pending}
      className={`${base} ${style} ${pending ? 'opacity-60' : ''}`}
    >
      {pending ? '…' : label}
    </button>
  )
}

// ── Day cell popover (PL + SI per breakdown) ──────────────────
function DayCellPopover({ cell, onClose, onFlip, anchorRect }) {
  const wrapRef = useRef(null)
  useEffect(() => {
    const handler = (e) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) onClose()
    }
    const esc = (e) => { if (e.key === 'Escape') onClose() }
    document.addEventListener('mousedown', handler)
    document.addEventListener('keydown', esc)
    return () => {
      document.removeEventListener('mousedown', handler)
      document.removeEventListener('keydown', esc)
    }
  }, [onClose])

  return (
    <div
      ref={wrapRef}
      className="fixed inset-0 z-40 flex items-center justify-center p-4"
      style={{ backgroundColor: 'rgba(0,0,0,0.4)' }}
      onClick={(e) => { if (e.target === e.currentTarget) onClose() }}
    >
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg p-5 space-y-3">
        <div className="flex items-center justify-between">
          <h3 className="text-base font-black text-navy">{fmtShortDate(cell.date)}</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600 text-2xl leading-none">×</button>
        </div>
        <div className="space-y-3 max-h-[60vh] overflow-y-auto">
          {(cell.breakdown || []).map((b, i) => (
            <div key={i} className="border border-slate-100 rounded-lg p-3 space-y-2">
              <div className="flex justify-between items-baseline">
                <p className="font-semibold text-sm text-slate-800">{b.contractor}</p>
                <p className="text-xs text-slate-500 font-mono">{b.contract_num} · {b.borough}</p>
              </div>
              {b.wo_ids?.length > 0 && (
                <p className="text-[11px] text-slate-500 font-mono">
                  WOs: {b.wo_ids.join(', ')}
                </p>
              )}
              <div className="flex items-center gap-3 flex-wrap">
                <div className="flex items-center gap-1.5">
                  <span className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">PL</span>
                  <TogglePill
                    label="Done"
                    on={b.pl.done}
                    onClick={() => onFlip(b.pl.doc_id, 'done', !b.pl.done)}
                  />
                  <TogglePill
                    label="Sent"
                    on={b.pl.sent}
                    color="green"
                    disabled={!b.pl.done}
                    onClick={() => onFlip(b.pl.doc_id, 'sent', !b.pl.sent)}
                  />
                </div>
                <div className="flex items-center gap-1.5">
                  <span className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">SI</span>
                  <TogglePill
                    label="Done"
                    on={b.si.done}
                    onClick={() => onFlip(b.si.doc_id, 'done', !b.si.done)}
                  />
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  )
}

// ── Week cell popover (CP per breakdown) ──────────────────────
function WeekCellPopover({ cell, onClose, onFlip }) {
  const wrapRef = useRef(null)
  useEffect(() => {
    const handler = (e) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) onClose()
    }
    const esc = (e) => { if (e.key === 'Escape') onClose() }
    document.addEventListener('mousedown', handler)
    document.addEventListener('keydown', esc)
    return () => {
      document.removeEventListener('mousedown', handler)
      document.removeEventListener('keydown', esc)
    }
  }, [onClose])

  return (
    <div
      ref={wrapRef}
      className="fixed inset-0 z-40 flex items-center justify-center p-4"
      style={{ backgroundColor: 'rgba(0,0,0,0.4)' }}
      onClick={(e) => { if (e.target === e.currentTarget) onClose() }}
    >
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg p-5 space-y-3">
        <div className="flex items-center justify-between">
          <h3 className="text-base font-black text-navy">Week of {fmtShortDate(cell.week_start)}</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600 text-2xl leading-none">×</button>
        </div>
        <div className="space-y-3 max-h-[60vh] overflow-y-auto">
          {(cell.breakdown || []).map((b, i) => (
            <div key={i} className="border border-slate-100 rounded-lg p-3 space-y-2">
              <div className="flex justify-between items-baseline">
                <p className="font-semibold text-sm text-slate-800">{b.contractor}</p>
                <p className="text-xs text-slate-500 font-mono">{b.contract_num} · {b.borough}</p>
              </div>
              {b.wo_ids?.length > 0 && (
                <p className="text-[11px] text-slate-500 font-mono">
                  WOs: {b.wo_ids.join(', ')}
                </p>
              )}
              <div className="flex items-center gap-1.5">
                <span className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">CP</span>
                <TogglePill
                  label="Done"
                  on={b.cp.done}
                  onClick={() => onFlip(b.cp.doc_id, 'done', !b.cp.done)}
                />
                <TogglePill
                  label="Sent"
                  on={b.cp.sent}
                  color="green"
                  disabled={!b.cp.done}
                  onClick={() => onFlip(b.cp.doc_id, 'sent', !b.cp.sent)}
                />
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  )
}

// ── Day calendar grid ─────────────────────────────────────────
function DayCalendar({ monthIso, days, onCellClick }) {
  const todayLocal = todayIso()
  const dayByDate = useMemo(() => {
    const m = {}
    ;(days || []).forEach(d => { m[d.date] = d })
    return m
  }, [days])

  const [y, m] = monthIso.split('-').map(Number)
  const firstDow = new Date(y, m - 1, 1).getDay()  // 0 = Sun
  const daysInMonth = new Date(y, m, 0).getDate()

  const cells = []
  // Leading blanks
  for (let i = 0; i < firstDow; i++) cells.push({ blank: true, key: 'lead-' + i })
  for (let d = 1; d <= daysInMonth; d++) {
    const iso = `${monthIso}-${String(d).padStart(2, '0')}`
    cells.push({ blank: false, key: iso, date: iso, day: d, cell: dayByDate[iso] })
  }
  // Trailing blanks to fill last row
  while (cells.length % 7 !== 0) cells.push({ blank: true, key: 'trail-' + cells.length })

  return (
    <div className="card p-4">
      <div className="grid grid-cols-7 gap-1 text-[10px] font-extrabold uppercase tracking-widest text-slate-400 mb-1">
        {['S','M','T','W','T','F','S'].map((d, i) => <div key={i} className="text-center">{d}</div>)}
      </div>
      <div className="grid grid-cols-7 gap-1">
        {cells.map(c => {
          if (c.blank) return <div key={c.key} className="aspect-square" />
          const cell = c.cell
          const status = cell?.status || ''
          const bg = status ? STATUS_BG[status] : 'bg-white text-slate-400 border-slate-100'
          const isToday = c.date === todayLocal
          const clickable = !!cell
          return (
            <button
              key={c.key}
              type="button"
              disabled={!clickable}
              onClick={() => clickable && onCellClick(cell)}
              className={`aspect-square rounded-md border text-xs font-semibold relative
                          flex items-start justify-end p-1.5 transition-all
                          ${bg}
                          ${isToday ? 'ring-2 ring-navy/40' : ''}
                          ${clickable ? 'cursor-pointer hover:ring-2 hover:ring-navy/40' : 'cursor-default'}`}
            >
              <span className="text-[11px]">{c.day}</span>
            </button>
          )
        })}
      </div>
    </div>
  )
}

// ── Week calendar (Sunday-anchored, ~5 cells per month) ───────
function WeekCalendar({ monthIso, weeks, onCellClick }) {
  const weekByStart = useMemo(() => {
    const m = {}
    ;(weeks || []).forEach(w => { m[w.week_start] = w })
    return m
  }, [weeks])

  // Build all weeks whose start falls in this month.
  const [y, m] = monthIso.split('-').map(Number)
  const firstDay = new Date(y, m - 1, 1)
  const lastDay  = new Date(y, m, 0)
  const cells = []
  // Find first Sunday on or before first of month
  const firstSun = new Date(firstDay)
  firstSun.setDate(firstDay.getDate() - firstDay.getDay())
  let cursor = new Date(firstSun)
  while (cursor <= lastDay) {
    const iso = `${cursor.getFullYear()}-${String(cursor.getMonth()+1).padStart(2,'0')}-${String(cursor.getDate()).padStart(2,'0')}`
    cells.push({ week_start: iso, cell: weekByStart[iso] })
    cursor.setDate(cursor.getDate() + 7)
  }

  return (
    <div className="card p-4">
      <p className="text-[10px] font-extrabold uppercase tracking-widest text-slate-400 mb-2">Week starting</p>
      <div className="space-y-1.5">
        {cells.map(c => {
          const cell = c.cell
          const status = cell?.status || ''
          const bg = status ? STATUS_BG[status] : 'bg-white text-slate-400 border-slate-100'
          const clickable = !!cell
          return (
            <button
              key={c.week_start}
              type="button"
              disabled={!clickable}
              onClick={() => clickable && onCellClick(cell)}
              className={`w-full rounded-md border text-sm font-semibold
                          flex items-center justify-between px-3 py-2 transition-all
                          ${bg}
                          ${clickable ? 'cursor-pointer hover:ring-2 hover:ring-navy/40' : 'cursor-default'}`}
            >
              <span>{fmtShortDate(c.week_start)}</span>
              {!status && <span className="text-[10px] text-slate-400 italic">no work</span>}
            </button>
          )
        })}
      </div>
    </div>
  )
}

// ── Pending list ──────────────────────────────────────────────
function PendingList({ kind, pending, onMark }) {
  const items = (pending || []).filter(p => p.kind === kind)
  if (items.length === 0) {
    return (
      <div className="card p-3 text-xs text-slate-400 italic">
        Nothing pending. ✓
      </div>
    )
  }
  return (
    <div className="card overflow-hidden">
      <div className="px-3 py-2 border-b border-slate-100 flex justify-between">
        <p className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">
          Pending ({items.length})
        </p>
        <p className="text-[10px] text-slate-400 italic">FIFO · oldest first</p>
      </div>
      <div className="max-h-[280px] overflow-y-auto divide-y divide-slate-100">
        {items.map((it, i) => (
          <div key={i} className="px-3 py-2 text-xs flex items-center justify-between gap-3">
            <div className="min-w-0 flex-1">
              <p className="font-mono text-slate-700">{it.anchor}</p>
              <p className="text-slate-500 truncate">
                {it.contractor} · <span className="font-mono">{it.contract_num}</span> · {it.borough}
              </p>
              <p className="text-[10px] text-amber-700 font-semibold">
                {it.missing.join(', ')}
              </p>
            </div>
            <div className="flex gap-1.5">
              {it.missing.includes('PL Done') && (
                <button onClick={() => onMark(it.doc_id, 'done', true)}
                  className="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded border bg-amber-500 text-white border-amber-500 hover:bg-amber-600">
                  Mark Done
                </button>
              )}
              {it.missing.includes('PL Sent') && (
                <button onClick={() => onMark(it.doc_id, 'sent', true)}
                  className="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded border bg-green-500 text-white border-green-500 hover:bg-green-600">
                  Mark Sent
                </button>
              )}
              {it.missing.includes('SI Done') && (
                <button onClick={() => onMark(it.doc_id, 'done', true)}
                  className="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded border bg-amber-500 text-white border-amber-500 hover:bg-amber-600">
                  Mark Done
                </button>
              )}
              {it.missing.includes('CP Done') && (
                <button onClick={() => onMark(it.doc_id, 'done', true)}
                  className="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded border bg-amber-500 text-white border-amber-500 hover:bg-amber-600">
                  Mark Done
                </button>
              )}
              {it.missing.includes('CP Sent') && (
                <button onClick={() => onMark(it.doc_id, 'sent', true)}
                  className="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded border bg-green-500 text-white border-green-500 hover:bg-green-600">
                  Mark Sent
                </button>
              )}
            </div>
          </div>
        ))}
      </div>
    </div>
  )
}

// ── Main ──────────────────────────────────────────────────────
export default function DocStatusTab() {
  const [monthIso, setMonthIso] = useState(todayMonthIso())
  const [data,    setData]      = useState(null)
  const [loading, setLoading]   = useState(true)
  const [error,   setError]     = useState(null)

  const [activeDay,  setActiveDay]  = useState(null)
  const [activeWeek, setActiveWeek] = useState(null)

  const load = async (iso) => {
    setLoading(true)
    setError(null)
    try {
      const d = await fetchDocStatus(iso)
      setData(d)
    } catch (e) {
      setError(e.message)
    } finally {
      setLoading(false)
    }
  }
  useEffect(() => { load(monthIso) }, [monthIso])

  // Optimistic flip + revert on failure
  const flip = async (docId, flag, value) => {
    // Optimistically mutate the local data.
    setData(prev => {
      if (!prev) return prev
      const updateInBreakdown = (b, key) => {
        if (!b[key] || b[key].doc_id !== docId) return b
        return { ...b, [key]: { ...b[key], [flag]: value } }
      }
      const days = prev.days?.map(d => ({
        ...d,
        breakdown: d.breakdown.map(b => updateInBreakdown(updateInBreakdown(b, 'pl'), 'si')),
      }))
      const weeks = prev.weeks?.map(w => ({
        ...w,
        breakdown: w.breakdown.map(b => updateInBreakdown(b, 'cp')),
      }))
      // Recompute cell statuses (matches server-side rollup)
      const recolorDay = (cell) => {
        let allFull = cell.breakdown.length > 0, allEmpty = true
        cell.breakdown.forEach(b => {
          const full  = b.pl.done && b.pl.sent && b.si.done
          const empty = !b.pl.done && !b.pl.sent && !b.si.done
          if (!full)  allFull  = false
          if (!empty) allEmpty = false
        })
        return { ...cell, status: allFull ? 'green' : (allEmpty ? 'gray' : 'amber') }
      }
      const recolorWeek = (cell) => {
        let allFull = cell.breakdown.length > 0, allEmpty = true
        cell.breakdown.forEach(b => {
          const full  = b.cp.done && b.cp.sent
          const empty = !b.cp.done && !b.cp.sent
          if (!full)  allFull  = false
          if (!empty) allEmpty = false
        })
        return { ...cell, status: allFull ? 'green' : (allEmpty ? 'gray' : 'amber') }
      }
      return {
        ...prev,
        days:  days?.map(recolorDay),
        weeks: weeks?.map(recolorWeek),
      }
    })

    try {
      await flipDocFlags([{ doc_id: docId, [flag]: value }])
      // Refetch in the background to pick up server-canonical state +
      // updated pending list. Doesn't block the optimistic UI.
      load(monthIso)
    } catch (e) {
      console.warn('flip failed, reverting:', e.message)
      load(monthIso)   // revert by hard refetch
    }
  }

  const onPendingMark = (docId, flag, value) => flip(docId, flag, value)

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div className="flex items-center gap-2">
          <button
            onClick={() => setMonthIso(shiftMonth(monthIso, -1))}
            className="btn-ghost text-sm px-3"
          >‹ Prev</button>
          <p className="text-lg font-black text-navy">{fmtMonth(monthIso)}</p>
          <button
            onClick={() => setMonthIso(shiftMonth(monthIso, +1))}
            className="btn-ghost text-sm px-3"
          >Next ›</button>
          {monthIso !== todayMonthIso() && (
            <button
              onClick={() => setMonthIso(todayMonthIso())}
              className="btn-ghost text-sm px-3"
            >Today</button>
          )}
        </div>
        {loading && <span className="text-xs text-slate-400">loading…</span>}
        {error && <span className="text-xs text-red-600">⚠ {error}</span>}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <div className="space-y-3">
          <p className="section-label">Production Logs + Sign-Ins</p>
          <DayCalendar monthIso={monthIso} days={data?.days} onCellClick={setActiveDay} />
          <PendingList kind="day" pending={data?.pending} onMark={onPendingMark} />
        </div>
        <div className="space-y-3">
          <p className="section-label">Certified Payroll</p>
          <WeekCalendar monthIso={monthIso} weeks={data?.weeks} onCellClick={setActiveWeek} />
          <PendingList kind="week" pending={data?.pending} onMark={onPendingMark} />
        </div>
      </div>

      {activeDay && (
        <DayCellPopover
          cell={activeDay}
          onClose={() => setActiveDay(null)}
          onFlip={flip}
        />
      )}
      {activeWeek && (
        <WeekCellPopover
          cell={activeWeek}
          onClose={() => setActiveWeek(null)}
          onFlip={flip}
        />
      )}
    </div>
  )
}
