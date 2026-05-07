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

// "Apr 13 – Apr 19" for a Sunday-anchored week-start ISO.
const fmtWeekRange = (sundayIso) => {
  const [y, m, d] = String(sundayIso || '').split('-').map(Number)
  if (!y || !m || !d) return sundayIso
  const start = new Date(y, m - 1, d)
  const end   = new Date(y, m - 1, d + 6)
  const a = `${MONTHS[start.getMonth()].slice(0, 3)} ${start.getDate()}`
  const b = `${MONTHS[end.getMonth()].slice(0, 3)} ${end.getDate()}`
  return `${a} – ${b}`
}

// ── Cell background per status ────────────────────────────────
const STATUS_BG = {
  gray:  'bg-slate-100  text-slate-500 border-slate-200',
  amber: 'bg-amber-100  text-amber-800 border-amber-200',
  green: 'bg-green-100  text-green-800 border-green-200',
}

// ── Pending-list action label per missing-flag ────────────────
const ACTION_TITLE = {
  'PL Done': 'Complete Production Log',
  'PL Sent': 'Send Production Log',
  'SI Done': 'Complete Sign-In Sheet',
  'CP Done': 'Complete Certified Payroll',
  'CP Sent': 'Send Certified Payroll',
}

// ── Bullet summary for a calendar cell's breakdown ────────────
//
// State icons:
//   ✓  all complete                    (green)
//   ◐  partial — append (n/total)      (amber)
//   ○  none yet                        (slate)
//   –  no work that day/week           (slate)
function bulletFor(label, n, total) {
  if (total === 0)  return { icon: '–', label,                              tone: 'slate' }
  if (n === 0)      return { icon: '○', label,                              tone: 'slate' }
  if (n === total)  return { icon: '✓', label,                              tone: 'green' }
  return                   { icon: '◐', label: `${label} ${n}/${total}`,    tone: 'amber' }
}

function summarizeBreakdown(breakdown, kind) {
  if (!breakdown || breakdown.length === 0) {
    return [{ icon: '–', label: 'No Work', tone: 'slate' }]
  }
  const total = breakdown.length
  if (kind === 'day') {
    // Order: SI first (sign-in must be done before a PL can exist),
    // then PL Done, then PL Sent.
    const siDone = breakdown.filter(b => b.si.done).length
    const plDone = breakdown.filter(b => b.pl.done).length
    const plSent = breakdown.filter(b => b.pl.sent).length
    return [
      bulletFor('SI Done', siDone, total),
      bulletFor('PL Done', plDone, total),
      bulletFor('PL Sent', plSent, total),
    ]
  }
  // kind === 'week' → CP only
  const cpDone = breakdown.filter(b => b.cp.done).length
  const cpSent = breakdown.filter(b => b.cp.sent).length
  return [
    bulletFor('CP Done', cpDone, total),
    bulletFor('CP Sent', cpSent, total),
  ]
}

function StatusBullet({ icon, label, tone }) {
  const text = {
    green: 'text-green-700',
    amber: 'text-amber-700',
    slate: 'text-slate-500',
  }[tone] || 'text-slate-500'
  return (
    <div className={`flex items-center gap-1 leading-tight ${text}`}>
      <span className="text-[10px] font-bold leading-none flex-shrink-0 w-2.5 text-center">{icon}</span>
      <span className="text-[9px] font-medium truncate">{label}</span>
    </div>
  )
}

// ── Loading overlay for calendars ─────────────────────────────
function CalendarLoadingOverlay() {
  return (
    <div className="card p-4 min-h-[280px] relative">
      <div className="absolute inset-0 flex flex-col items-center justify-center gap-2">
        <div className="w-6 h-6 rounded-full border-2 border-slate-300 border-t-navy animate-spin" />
        <span className="text-xs text-slate-500 font-medium">Loading…</span>
      </div>
    </div>
  )
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
                  <span className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">SI</span>
                  <TogglePill
                    label="Done"
                    on={b.si.done}
                    onClick={() => onFlip(b.si.doc_id, 'done', !b.si.done)}
                  />
                </div>
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
function DayCalendar({ monthIso, days, loading, onCellClick }) {
  const todayLocal = todayIso()
  const dayByDate = useMemo(() => {
    const m = {}
    ;(days || []).forEach(d => { m[d.date] = d })
    return m
  }, [days])

  // First load: no data yet → full skeleton + spinner.
  if (loading && !days) return <CalendarLoadingOverlay />

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
    <div className={`card p-4 ${loading ? 'opacity-60 transition-opacity' : ''}`}>
      <div className="grid grid-cols-7 gap-1 text-[10px] font-extrabold uppercase tracking-widest text-slate-400 mb-1">
        {['S','M','T','W','T','F','S'].map((d, i) => <div key={i} className="text-center">{d}</div>)}
      </div>
      <div className="grid grid-cols-7 gap-1">
        {cells.map(c => {
          if (c.blank) return <div key={c.key} className="min-h-[64px]" />
          const cell = c.cell
          const status = cell?.status || 'gray'
          const bg = STATUS_BG[status] || STATUS_BG.gray
          const isToday = c.date === todayLocal
          const clickable = !!cell
          const bullets = summarizeBreakdown(cell?.breakdown, 'day')
          return (
            <button
              key={c.key}
              type="button"
              disabled={!clickable}
              onClick={() => clickable && onCellClick(cell)}
              className={`min-h-[64px] rounded-md border text-left p-1.5 transition-all
                          flex flex-col items-stretch
                          ${bg}
                          ${isToday ? 'ring-2 ring-navy/40' : ''}
                          ${clickable ? 'cursor-pointer hover:ring-2 hover:ring-navy/40' : 'cursor-default'}`}
            >
              <span className="text-[11px] font-semibold text-slate-700 self-end">{c.day}</span>
              <div className="mt-1 space-y-0.5 w-full">
                {bullets.map((b, i) => (
                  <StatusBullet key={i} icon={b.icon} label={b.label} tone={b.tone} />
                ))}
              </div>
            </button>
          )
        })}
      </div>
    </div>
  )
}

// ── Week calendar (Sunday-anchored, ~5 cells per month) ───────
function WeekCalendar({ monthIso, weeks, loading, onCellClick }) {
  const weekByStart = useMemo(() => {
    const m = {}
    ;(weeks || []).forEach(w => { m[w.week_start] = w })
    return m
  }, [weeks])

  // First load: no data yet → full skeleton + spinner.
  if (loading && !weeks) return <CalendarLoadingOverlay />

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
    <div className={`card p-4 ${loading ? 'opacity-60 transition-opacity' : ''}`}>
      <p className="text-[10px] font-extrabold uppercase tracking-widest text-slate-400 mb-2">Week of</p>
      <div className="space-y-2">
        {cells.map(c => {
          const cell = c.cell
          const status = cell?.status || 'gray'
          const bg = STATUS_BG[status] || STATUS_BG.gray
          const clickable = !!cell
          const bullets = summarizeBreakdown(cell?.breakdown, 'week')
          return (
            <button
              key={c.week_start}
              type="button"
              disabled={!clickable}
              onClick={() => clickable && onCellClick(cell)}
              className={`w-full rounded-md border text-left p-2 transition-all
                          flex flex-col items-stretch min-h-[64px]
                          ${bg}
                          ${clickable ? 'cursor-pointer hover:ring-2 hover:ring-navy/40' : 'cursor-default'}`}
            >
              <span className="text-[11px] font-semibold text-slate-700">
                {fmtWeekRange(c.week_start)}
              </span>
              <div className="mt-1 space-y-0.5 w-full">
                {bullets.map((b, i) => (
                  <StatusBullet key={i} icon={b.icon} label={b.label} tone={b.tone} />
                ))}
              </div>
            </button>
          )
        })}
      </div>
    </div>
  )
}

// ── Pending list action button (per-button loading state) ─────
function PendingActionButton({ item, onMark }) {
  const [pending, setPending] = useState(false)
  const flag = item.missing[0]
  const isSent = flag === 'PL Sent' || flag === 'CP Sent'
  const label  = isSent ? 'Mark Sent' : 'Mark Done'

  const handle = async () => {
    if (pending) return
    setPending(true)
    try {
      await onMark(item.doc_id, isSent ? 'sent' : 'done', true)
    } finally {
      setPending(false)
    }
  }

  const base = 'text-[10px] font-bold uppercase tracking-wider px-2.5 py-1 rounded border transition-all flex-shrink-0 min-w-[80px] text-center'
  const style = isSent
    ? 'bg-green-500 text-white border-green-500 hover:bg-green-600'
    : 'bg-amber-500 text-white border-amber-500 hover:bg-amber-600'

  return (
    <button
      type="button"
      onClick={handle}
      disabled={pending}
      className={`${base} ${style} ${pending ? 'opacity-60 cursor-wait' : ''}`}
    >
      {pending ? '…' : label}
    </button>
  )
}

// ── Pending list ──────────────────────────────────────────────
function PendingList({ kind, pending, loading, onMark }) {
  if (loading && !pending) {
    return (
      <div className="card p-4 text-xs text-slate-400 italic flex items-center justify-center min-h-[80px]">
        Loading…
      </div>
    )
  }
  const items = (pending || []).filter(p => p.kind === kind)
  if (items.length === 0) {
    return (
      <div className="card p-3 text-xs text-slate-400 italic">
        Nothing pending. ✓
      </div>
    )
  }
  return (
    <div className={`card overflow-hidden ${loading ? 'opacity-60 transition-opacity' : ''}`}>
      <div className="px-3 py-2 border-b border-slate-100 flex justify-between">
        <p className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">
          Pending ({items.length})
        </p>
        <p className="text-[10px] text-slate-400 italic">FIFO · oldest first</p>
      </div>
      <div className="max-h-[280px] overflow-y-auto divide-y divide-slate-100">
        {items.map((it, i) => (
          <div key={i} className="px-3 py-2.5 flex items-center justify-between gap-3">
            <div className="min-w-0 flex-1">
              <p className="font-semibold text-sm text-slate-800 truncate">
                {ACTION_TITLE[it.missing[0]] || it.missing.join(', ')}
              </p>
              <p className="text-xs text-slate-500 mt-0.5">
                {it.kind === 'week' ? `Week of ${fmtShortDate(it.anchor)}` : fmtShortDate(it.anchor)}
              </p>
              <p className="text-xs text-slate-500 truncate">
                {it.contractor} · <span className="font-mono">{it.contract_num}</span> · {it.borough}
              </p>
            </div>
            <PendingActionButton item={it} onMark={onMark} />
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
        {error && <span className="text-xs text-red-600">⚠ {error}</span>}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <div className="space-y-3">
          <p className="section-label">Production Logs + Sign-Ins</p>
          <DayCalendar monthIso={monthIso} days={data?.days} loading={loading} onCellClick={setActiveDay} />
          <PendingList kind="day" pending={data?.pending} loading={loading} onMark={onPendingMark} />
        </div>
        <div className="space-y-3">
          <p className="section-label">Certified Payroll</p>
          <WeekCalendar monthIso={monthIso} weeks={data?.weeks} loading={loading} onCellClick={setActiveWeek} />
          <PendingList kind="week" pending={data?.pending} loading={loading} onMark={onPendingMark} />
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
