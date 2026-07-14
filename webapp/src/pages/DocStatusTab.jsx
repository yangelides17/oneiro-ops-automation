import { useEffect, useMemo, useRef, useState } from 'react'
import GenerateDocModal from '../components/GenerateDocModal'
import PaystubUpload from '../components/PaystubUpload'
import WODocsQueue from '../components/WODocsQueue'
import { usePendingCounts } from '../lib/PendingCountsContext'

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

// ── Month-end documents (per month, per contract-borough) ─────
// Mirrors MONTH_END_DOCS_ in Code.js — keep keys/labels in sync. The
// three certificates are tracked as one "Certificates" line item (they
// ship together); Employee Utilization stays separate. Completed once a
// month per (contract, borough) pair and folded into the last CP week.
const MONTH_END_DOCS = [
  { key: 'EU',   label: 'Employee Utilization', short: 'EU' },
  { key: 'CERT', label: 'Certificates',         short: 'CC',
    note: "Contractor's · Compliance · 220 Labor Law" },
]

const MONTH_END_KEYS = new Set(MONTH_END_DOCS.map(d => d.key))
const MONTH_END_NOTE = Object.fromEntries(MONTH_END_DOCS.map(d => [d.key, d.note || '']))

// Render the Employee Utilization counts as one transcribable line, in
// the printed form's column order using its own abbreviations:
//   "TOT 12 · B 3 · H 4 · FEM 2"
// Zero-count races are omitted to keep it short; TOT always shows (an
// all-White crew renders "TOT 5" alone, since the form has no White
// column). OTH covers registry gaps and never hides from the total.
function utilizationLine(u) {
  const parts = [`TOT ${u.total}`]
  ;(u.buckets || []).forEach(b => { if (b.count > 0) parts.push(`${b.key} ${b.count}`) })
  if (u.other > 0) parts.push(`OTH ${u.other}`)
  return parts.join(' · ')
}

// A month-end pending item is anchored to a whole month, not a week.
// Detect it from the missing flag's doc key and pull the month out of
// the doc_id (e.g. "EU_2026-06_84125MBTP701_BK" → "2026-06"). Returns
// the month string for such items, else null.
function monthEndPendingMonth(item) {
  const key = String(item?.missing?.[0] || '').split(' ')[0]
  if (!MONTH_END_KEYS.has(key)) return null
  const m = String(item?.doc_id || '').match(/_(\d{4}-\d{2})_/)
  return m ? m[1] : null
}

// ── Pending-list action label per missing-flag ────────────────
const ACTION_TITLE = {
  'PL Done': 'Complete Production Log',
  'PL Sent': 'Send Production Log',
  'SI Done': 'Complete Sign-In Sheet',
  'CP Done': 'Complete Certified Payroll',
  'CP Sent': 'Send Certified Payroll',
  // Month-end docs — one Complete/Send pair per doc.
  ...MONTH_END_DOCS.reduce((m, d) => {
    m[`${d.key} Done`] = `Complete ${d.label}`
    m[`${d.key} Sent`] = `Send ${d.label}`
    return m
  }, {}),
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
  if (kind === 'day') {
    // Day breakdown is contractor-grouped: PL state lives on the group
    // (one PL per contractor per day), SI lives on each contracts[]
    // sub-entry (one SI per (contract, borough)). PL counts use only
    // PL-eligible contractors; SI counts use every contract.
    const plRows = breakdown.filter(b => b.pl_required)
    const totalPlCtrs = plRows.length
    const plDone = plRows.filter(b => b.pl?.done).length
    const plSent = plRows.filter(b => b.pl?.sent).length
    let totalContracts = 0, siDone = 0
    breakdown.forEach(b => {
      const contracts = b.contracts || []
      totalContracts += contracts.length
      siDone += contracts.filter(c => c.si?.done).length
    })
    const bullets = [bulletFor('SI Done', siDone, totalContracts)]
    if (totalPlCtrs > 0) {
      bullets.push(bulletFor('PL Done', plDone, totalPlCtrs))
      bullets.push(bulletFor('PL Sent', plSent, totalPlCtrs))
    }
    return bullets
  }
  // kind === 'week' → CP only
  const total = breakdown.length
  const cpDone = breakdown.filter(b => b.cp.done).length
  const cpSent = breakdown.filter(b => b.cp.sent).length
  return [
    bulletFor('CP Done', cpDone, total),
    bulletFor('CP Sent', cpSent, total),
  ]
}

// ── Month-end cluster: bullets rolled up per doc across all pairs ──
// Mirrors the CP style — each of the four docs gets a Done and a Sent
// bullet (n = pairs with the flag set, total = pairs that worked the
// month). Returns [] when nothing worked the month.
function summarizeMonthEnd(breakdown) {
  const rows = breakdown || []
  if (rows.length === 0) return []
  const total = rows.length
  const docFor = (r, key) => (r.docs || []).find(m => m.key === key)
  return MONTH_END_DOCS.flatMap((d) => {
    const done = rows.filter(r => docFor(r, d.key)?.done).length
    const sent = rows.filter(r => docFor(r, d.key)?.sent).length
    return [
      bulletFor(`${d.short} Done`, done, total),
      bulletFor(`${d.short} Sent`, sent, total),
    ]
  })
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

// ── Toggle pill (binary: green=on, slate=off; amber reserved for rollups) ─
function TogglePill({ label, on, disabled, pending, onClick }) {
  const base = 'text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded border transition-all'
  let style
  if (disabled) {
    style = 'bg-slate-50 text-slate-300 border-slate-100 cursor-not-allowed'
  } else if (on) {
    style = 'bg-green-500 text-white border-green-500 hover:bg-green-600'
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

// ── Status rollup for a single contractor breakdown row ───────
//   day  → 'gray' | 'amber' | 'green' from contractor.pl (when required)
//                                      + every contracts[].si
//   week → 'gray' | 'amber' | 'green' from cp.done + cp.sent
function dayBreakdownStatus(b) {
  const contracts = b.contracts || []
  const siAllDone = contracts.length > 0 && contracts.every(c => c.si?.done)
  const siAllBlank = contracts.every(c => !c.si?.done)
  const plRequired = !!b.pl_required
  const plFull  = !plRequired || (b.pl?.done && b.pl?.sent)
  const plBlank = !plRequired || (!b.pl?.done && !b.pl?.sent)
  const allDone  = siAllDone && plFull
  const allBlank = siAllBlank && plBlank
  return allDone ? 'green' : allBlank ? 'gray' : 'amber'
}
function weekBreakdownStatus(b) {
  const flags = [b.cp?.done, b.cp?.sent]
  const allDone = flags.every(Boolean)
  const noneDone = flags.every(f => !f)
  return allDone ? 'green' : noneDone ? 'gray' : 'amber'
}

// Per-pair rollup for a month-end breakdown row (all four docs).
function monthEndRowStatus(row) {
  const flags = (row.docs || []).flatMap(m => [m.done, m.sent])
  const allDone = flags.length > 0 && flags.every(Boolean)
  const noneDone = flags.every(f => !f)
  return allDone ? 'green' : noneDone ? 'gray' : 'amber'
}

// Combined status for the month's last-week cell: rolls up BOTH the
// week's CP flags AND the month-end docs (which live on this cell).
// Used only for the last-week card, whose color reflects everything
// shown on it. Other week cells stay CP-only (cell.status).
function combinedLastWeekStatus(cell, monthEnd) {
  const flags = []
  ;(cell?.breakdown || []).forEach(b => { flags.push(!!b.cp?.done, !!b.cp?.sent) })
  ;(monthEnd?.breakdown || []).forEach(r => {
    ;(r.docs || []).forEach(m => flags.push(!!m.done, !!m.sent))
  })
  if (flags.length === 0) return 'gray'
  if (flags.every(Boolean)) return 'green'
  if (flags.every(f => !f)) return 'gray'
  return 'amber'
}

// ── Day cell popover (PL + SI per breakdown) ──────────────────
function DayCellPopover({ cell, onClose, onFlip, isPending, anchorRect }) {
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
          {(cell.breakdown || []).map((b, i) => {
            const rowBg = STATUS_BG[dayBreakdownStatus(b)]
            const contracts = b.contracts || []
            return (
              <div key={i} className={`border rounded-lg p-3 space-y-2 ${rowBg}`}>
                <div className="flex justify-between items-baseline">
                  <p className="font-semibold text-sm text-slate-800">{b.contractor}</p>
                  {b.crew_chief && (
                    <p className="text-[10px] text-slate-600">
                      <span className="font-semibold">Crew:</span> {b.crew_chief}
                    </p>
                  )}
                  {contracts.length > 1 && (
                    <p className="text-[10px] text-slate-500">
                      {contracts.length} contracts
                    </p>
                  )}
                </div>
                {/* One row per contract — SI (per-contract) + PL (per-contractor)
                    toggles all on the same line. PL toggles flip the contractor's
                    single PL doc_id no matter which row they're clicked from. */}
                <div className="space-y-1.5 pl-2 border-l-2 border-slate-300/60">
                  {contracts.map((c, j) => (
                    <div key={j} className="flex items-center justify-between gap-2 py-1">
                      <div className="min-w-0">
                        <p className="text-xs text-slate-700 font-mono truncate">
                          {c.contract_num} · {c.borough}
                        </p>
                        {c.wo_ids?.length > 0 && (
                          <p className="text-[10px] text-slate-500 font-mono truncate">
                            WOs: {c.wo_ids.join(', ')}
                          </p>
                        )}
                      </div>
                      <div className="flex items-center gap-2 flex-shrink-0">
                        <div className="flex items-center gap-1">
                          <span className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">SI</span>
                          <TogglePill
                            label="Done"
                            on={c.si.done}
                            pending={isPending(c.si.doc_id, 'done')}
                            onClick={() => onFlip(c.si.doc_id, 'done', !c.si.done)}
                          />
                        </div>
                        {b.pl_required && (
                          <div className="flex items-center gap-1">
                            <span className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">PL</span>
                            <TogglePill
                              label="Done"
                              on={b.pl.done}
                              pending={isPending(b.pl.doc_id, 'done')}
                              onClick={() => onFlip(b.pl.doc_id, 'done', !b.pl.done)}
                            />
                            <TogglePill
                              label="Sent"
                              on={b.pl.sent}
                              disabled={!b.pl.done}
                              pending={isPending(b.pl.doc_id, 'sent')}
                              onClick={() => onFlip(b.pl.doc_id, 'sent', !b.pl.sent)}
                            />
                          </div>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )
          })}
        </div>
      </div>
    </div>
  )
}

// ── Week cell popover (CP per breakdown) ──────────────────────
// On the month's last worked week (showMonthEnd), the CP entries render
// at the top and the month-end docs (one section per pair, all four
// docs) render underneath — everything in one popover.
function WeekCellPopover({ cell, monthEnd, showMonthEnd, onClose, onFlip, isPending }) {
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

  const meRows = showMonthEnd ? (monthEnd?.breakdown || []) : []

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
          {/* Certified Payroll — top */}
          <p className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">
            Certified Payroll
          </p>
          {(cell.breakdown || []).length === 0 && (
            <p className="text-xs text-slate-400 italic">No Certified Payroll this week.</p>
          )}
          {(cell.breakdown || []).map((b, i) => {
            const rowBg = STATUS_BG[weekBreakdownStatus(b)]
            return (
              <div key={i} className={`border rounded-lg p-3 space-y-2 ${rowBg}`}>
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
                    pending={isPending(b.cp.doc_id, 'done')}
                    onClick={() => onFlip(b.cp.doc_id, 'done', !b.cp.done)}
                  />
                  <TogglePill
                    label="Sent"
                    on={b.cp.sent}
                    disabled={!b.cp.done}
                    pending={isPending(b.cp.doc_id, 'sent')}
                    onClick={() => onFlip(b.cp.doc_id, 'sent', !b.cp.sent)}
                  />
                </div>
              </div>
            )
          })}

          {/* Month-End Documents — underneath, one section per pair */}
          {meRows.length > 0 && (
            <div className="pt-2 border-t-2 border-slate-200 space-y-3">
              <p className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">
                Month-End Documents · {fmtMonth(monthEnd.month)}
              </p>
              {meRows.map((row, i) => {
                const rowBg = STATUS_BG[monthEndRowStatus(row)]
                return (
                  <div key={i} className={`border rounded-lg p-3 space-y-2 ${rowBg}`}>
                    <div className="flex justify-between items-baseline">
                      <p className="font-semibold text-sm text-slate-800">{row.contractor}</p>
                      <p className="text-xs text-slate-500 font-mono">{row.contract_num} · {row.borough}</p>
                    </div>
                    {row.contract_id && (
                      <p className="text-[11px] text-slate-500">
                        <span className="font-semibold">Contract ID:</span> <span className="font-mono">{row.contract_id}</span>
                      </p>
                    )}
                    {row.utilization && (
                      <div className="space-y-0.5">
                        <p className="text-[11px] text-slate-500">
                          <span className="font-semibold">Employee Utilization:</span>{' '}
                          <span className="font-mono">{utilizationLine(row.utilization)}</span>
                        </p>
                        {(row.utilization.warnings || []).map((w, wi) => (
                          <p key={wi} className="text-[10px] text-amber-600">⚠ {w}</p>
                        ))}
                      </div>
                    )}
                    <div className="space-y-1.5 pl-2 border-l-2 border-slate-300/60">
                      {(row.docs || []).map((mdoc) => (
                        <div key={mdoc.key} className="flex items-center justify-between gap-2">
                          <div className="min-w-0">
                            <span className="text-xs text-slate-700 truncate block">{mdoc.label}</span>
                            {MONTH_END_NOTE[mdoc.key] && (
                              <span className="text-[10px] text-slate-400 block truncate">{MONTH_END_NOTE[mdoc.key]}</span>
                            )}
                          </div>
                          <div className="flex items-center gap-1.5 flex-shrink-0">
                            <TogglePill
                              label="Done"
                              on={mdoc.done}
                              pending={isPending(mdoc.doc_id, 'done')}
                              onClick={() => onFlip(mdoc.doc_id, 'done', !mdoc.done)}
                            />
                            <TogglePill
                              label="Sent"
                              on={mdoc.sent}
                              disabled={!mdoc.done}
                              pending={isPending(mdoc.doc_id, 'sent')}
                              onClick={() => onFlip(mdoc.doc_id, 'sent', !mdoc.sent)}
                            />
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )
              })}
            </div>
          )}
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
          if (c.blank) return <div key={c.key} className="h-[80px]" />
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
              onClick={() => clickable && onCellClick(c.date)}
              className={`h-[80px] rounded-md border text-left p-1.5 transition-all
                          flex flex-col items-stretch overflow-hidden
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
// The month's last worked week additionally carries the month-end docs:
// CP bullets stay left-justified; the 8 month-end bullets sit to the
// right in two columns, and that cell's color rolls up both.
function WeekCalendar({ monthIso, weeks, monthEnd, lastWeekStart, loading, onCellClick }) {
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

  const meBullets = summarizeMonthEnd(monthEnd?.breakdown)

  return (
    <div className={`card p-4 ${loading ? 'opacity-60 transition-opacity' : ''}`}>
      <p className="text-[10px] font-extrabold uppercase tracking-widest text-slate-400 mb-1">Week of</p>
      <div className="space-y-1">
        {cells.map(c => {
          const cell = c.cell
          // The last calendar week carries the month-end docs — even when
          // it had no CP work. Then it's still clickable (opens the
          // month-end popover) and shows only the month-end bullets.
          const isLastWeek = c.week_start === lastWeekStart && meBullets.length > 0
          const clickable = !!cell || isLastWeek
          const effectiveCell = cell || (isLastWeek ? { week_start: c.week_start, breakdown: [] } : null)
          const status = isLastWeek
            ? combinedLastWeekStatus(cell, monthEnd)
            : (cell?.status || 'gray')
          const bg = STATUS_BG[status] || STATUS_BG.gray
          // CP bullets always render (even "– No Work" on the last week
          // when that week had no CP) so it's clear why there's no CP.
          const cpBullets = summarizeBreakdown(cell?.breakdown, 'week')
          return (
            <button
              key={c.week_start}
              type="button"
              disabled={!clickable}
              onClick={() => clickable && onCellClick(c.week_start)}
              className={`w-full rounded-md border text-left transition-all
                          flex flex-col items-stretch overflow-hidden
                          ${isLastWeek ? 'min-h-[80px] p-2.5' : 'h-[80px] p-1.5'}
                          ${bg}
                          ${clickable ? 'cursor-pointer hover:ring-2 hover:ring-navy/40' : 'cursor-default'}`}
            >
              <span className="text-[11px] font-semibold text-slate-700">
                {fmtWeekRange(c.week_start)}
              </span>
              {isLastWeek ? (
                <div className="mt-1.5 flex gap-8 w-full">
                  {/* CP bullets — left-justified, always shown */}
                  <div className="space-y-1 flex-shrink-0">
                    {cpBullets.map((b, i) => (
                      <StatusBullet key={i} icon={b.icon} label={b.label} tone={b.tone} />
                    ))}
                  </div>
                  {/* Month-end bullets — single column (EU + Certificates) */}
                  <div className="space-y-1 flex-1 min-w-0">
                    {meBullets.map((b, i) => (
                      <StatusBullet key={i} icon={b.icon} label={b.label} tone={b.tone} />
                    ))}
                  </div>
                </div>
              ) : (
                <div className="mt-1 space-y-0.5 w-full">
                  {cpBullets.map((b, i) => (
                    <StatusBullet key={i} icon={b.icon} label={b.label} tone={b.tone} />
                  ))}
                </div>
              )}
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
  const isSent = / Sent$/.test(flag)
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

  // Both Mark Done and Mark Sent share the green palette — completion =
  // green is the convention. Slate is reserved for "pending / not done"
  // and amber is reserved for partial-rollup states.
  const base = 'text-[10px] font-bold uppercase tracking-wider px-2.5 py-1 rounded border transition-all flex-shrink-0 min-w-[80px] text-center'
  const style = 'bg-green-500 text-white border-green-500 hover:bg-green-600'

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

// ── Generate button (for PL Done / CP Done pending items) ─────
function PendingGenerateButton({ item, onGenerate }) {
  const handle = () => onGenerate?.(item)
  return (
    <button
      type="button"
      onClick={handle}
      className="text-[10px] font-bold uppercase tracking-wider px-2.5 py-1 rounded border transition-all flex-shrink-0 min-w-[80px] text-center bg-green-500 text-white border-green-500 hover:bg-green-600"
    >
      Generate
    </button>
  )
}

const canGenerateForItem = (it) =>
  it && (it.missing?.[0] === 'PL Done' || it.missing?.[0] === 'CP Done')

// ── Pending list ──────────────────────────────────────────────
function PendingList({ kind, pending, loading, onMark, onGenerate }) {
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
        {items.map((it, i) => {
          const meMonth = monthEndPendingMonth(it)
          return (
          <div key={i} className="px-3 py-2.5 flex items-center justify-between gap-3">
            <div className="min-w-0 flex-1">
              <p className="font-semibold text-sm text-slate-800 truncate">
                {ACTION_TITLE[it.missing[0]] || it.missing.join(', ')}
              </p>
              <p className="text-xs text-slate-500 mt-0.5">
                {meMonth
                  ? `Month of ${fmtMonth(meMonth)}`
                  : (it.kind === 'week' ? `Week of ${fmtShortDate(it.anchor)}` : fmtShortDate(it.anchor))}
              </p>
              <p className="text-xs text-slate-500 truncate">
                {it.contract_num
                  ? <>{it.contractor} · <span className="font-mono">{it.contract_num}</span> · {it.borough}</>
                  : it.contractor}
                {it.crew_chief && (
                  <> · <span className="font-semibold text-slate-600">{it.crew_chief}</span></>
                )}
              </p>
            </div>
            <div className="flex gap-1.5 flex-shrink-0 items-center">
              {canGenerateForItem(it) && (
                <PendingGenerateButton item={it} onGenerate={onGenerate} />
              )}
              <PendingActionButton item={it} onMark={onMark} />
            </div>
          </div>
          )
        })}
      </div>
    </div>
  )
}

// ── Main ──────────────────────────────────────────────────────
export default function DocStatusTab({ wos, qbConnected, onDocsChange, onInvoiced }) {
  const [monthIso, setMonthIso] = useState(todayMonthIso())
  const [data,    setData]      = useState(null)
  const [loading, setLoading]   = useState(true)
  const [error,   setError]     = useState(null)

  // Popover targets hold the cell KEY (date / week_start), not a snapshot —
  // the live cell is derived from `data` at render so optimistic flips show
  // in the open modal without a close/reopen.
  const [activeDay,  setActiveDay]  = useState(null)
  const [activeWeek, setActiveWeek] = useState(null)

  // Per-pill in-flight guard (Set of `${docId}|${flag}`): each pill shows its
  // own "…" while its request runs; different pills/docs flip in parallel.
  // inFlightRef coalesces the post-burst reconcile; flipError surfaces a
  // failed save so nothing is dropped unnoticed.
  const [pendingFlags, setPendingFlags] = useState(() => new Set())
  const inFlightRef = useRef(0)
  const [flipError, setFlipError] = useState('')
  const monthIsoRef = useRef(monthIso)
  useEffect(() => { monthIsoRef.current = monthIso }, [monthIso])
  useEffect(() => {
    if (!flipError) return
    const t = setTimeout(() => setFlipError(''), 4000)
    return () => clearTimeout(t)
  }, [flipError])

  // GenerateDocModal state — `pendingItem` is the pending-list row the
  // user clicked Generate on; null = modal closed.
  const [pendingItem, setPendingItem] = useState(null)
  // Parsed paystub rows for the CP generate modal (null = none uploaded).
  const [paystub, setPaystub] = useState(null)

  // Push the pending count into the shared context so the Doc Status
  // tab in Dashboard's TabStrip can show a badge. Sourced via useEffect
  // on `data` (not from inside load) so the flip path — which mutates
  // data optimistically + triggers a background refetch — also keeps
  // the badge in sync.
  const { setCount } = usePendingCounts()

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

  // Keep the Doc Status nav badge in sync with whatever's currently in
  // data.pending — covers load (mount + month change) AND the flip
  // handler's background refetch.
  useEffect(() => {
    if (data == null) return
    setCount('doc_status_pending', (data.pending || []).length)
  }, [data, setCount])

  // The month's last CALENDAR week (Sunday of the week containing the
  // month's last day) — the card that carries the month-end docs, even
  // if that week had no CP work.
  const lastWeekStart = useMemo(() => {
    const [y, m] = monthIso.split('-').map(Number)
    const lastDay = new Date(y, m, 0)
    const sun = new Date(lastDay)
    sun.setDate(lastDay.getDate() - lastDay.getDay())
    return `${sun.getFullYear()}-${String(sun.getMonth() + 1).padStart(2, '0')}-${String(sun.getDate()).padStart(2, '0')}`
  }, [monthIso])

  // Apply a done/sent flip to `data` immutably — used for both the
  // optimistic update and the revert on failure.
  const applyFlag = (prev, docId, flag, value) => {
      if (!prev) return prev
      // Set the flag on a doc; clearing DONE also clears SENT (Sent implies
      // Done), matching the server cascade + DocStatusChips behavior.
      const applyTo = (obj) => {
        const next = { ...obj, [flag]: value }
        if (flag === 'done' && value === false && next.sent) next.sent = false
        return next
      }
      // Day breakdown: flip PL on the contractor row OR SI on any matching
      // contract sub-entry. Either match-or-pass-through.
      const updateDayBreakdown = (b) => {
        let next = b
        if (b.pl && b.pl.doc_id === docId) {
          next = { ...next, pl: applyTo(next.pl) }
        }
        if (Array.isArray(b.contracts)) {
          let touched = false
          const contracts = b.contracts.map(c => {
            if (c.si && c.si.doc_id === docId) {
              touched = true
              return { ...c, si: applyTo(c.si) }
            }
            return c
          })
          if (touched) next = { ...next, contracts }
        }
        return next
      }
      const updateWeekBreakdown = (b) => {
        if (!b.cp || b.cp.doc_id !== docId) return b
        return { ...b, cp: applyTo(b.cp) }
      }
      // Month-end docs: flip the matching doc on any pair row.
      const updateMonthEndRow = (row) => {
        if (!Array.isArray(row.docs)) return row
        let touched = false
        const docs = row.docs.map(m => {
          if (m.doc_id === docId) { touched = true; return applyTo(m) }
          return m
        })
        return touched ? { ...row, docs } : row
      }
      const days = prev.days?.map(d => ({
        ...d,
        breakdown: d.breakdown.map(updateDayBreakdown),
      }))
      const weeks = prev.weeks?.map(w => ({
        ...w,
        breakdown: w.breakdown.map(updateWeekBreakdown),
      }))
      const recolorDay = (cell) => {
        let allFull = cell.breakdown.length > 0, allEmpty = true
        cell.breakdown.forEach(b => {
          const contracts = b.contracts || []
          const siAllDone  = contracts.length > 0 && contracts.every(c => c.si?.done)
          const siAllBlank = contracts.every(c => !c.si?.done)
          const plRequired = !!b.pl_required
          const plFull  = !plRequired || (b.pl?.done && b.pl?.sent)
          const plBlank = !plRequired || (!b.pl?.done && !b.pl?.sent)
          const full  = plFull  && siAllDone
          const empty = plBlank && siAllBlank
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
      // Month-end cluster: flip the matching doc, then recolor.
      let month_end = prev.month_end
      if (month_end?.breakdown) {
        const breakdown = month_end.breakdown.map(updateMonthEndRow)
        let allFull = breakdown.length > 0, allEmpty = true
        breakdown.forEach(row => {
          ;(row.docs || []).forEach(m => {
            const full  = m.done && m.sent
            const empty = !m.done && !m.sent
            if (!full)  allFull  = false
            if (!empty) allEmpty = false
          })
        })
        month_end = { ...month_end, breakdown, status: allFull ? 'green' : (allEmpty ? 'gray' : 'amber') }
      }
      return {
        ...prev,
        days:  days?.map(recolorDay),
        weeks: weeks?.map(recolorWeek),
        month_end,
      }
  }

  // Silent reconcile after a burst of flips settles: refresh the pending
  // list + canonical state without the loading dim, and without clobbering
  // an optimistic change that's still in flight.
  const reconcile = async () => {
    const iso = monthIsoRef.current
    try {
      const d = await fetchDocStatus(iso)
      if (inFlightRef.current === 0 && iso === monthIsoRef.current) setData(d)
    } catch (e) {
      // Keep optimistic state; a later flip/reconcile or month load corrects it.
    }
  }

  // Optimistic pill flip. Per-pill in-flight guard lets different pills/docs
  // flip in parallel; the server serializes writes (LockService) so nothing
  // duplicates. Reverts + surfaces a visible error on failure.
  const flip = async (docId, flag, value) => {
    const key = docId + '|' + flag
    if (pendingFlags.has(key)) return
    setData(prev => applyFlag(prev, docId, flag, value))
    setPendingFlags(prev => { const n = new Set(prev); n.add(key); return n })
    inFlightRef.current++
    try {
      await flipDocFlags([{ doc_id: docId, [flag]: value }])
    } catch (e) {
      setData(prev => applyFlag(prev, docId, flag, !value))
      setFlipError("Couldn't save that change — please retry.")
    } finally {
      setPendingFlags(prev => { const n = new Set(prev); n.delete(key); return n })
      if (--inFlightRef.current === 0) reconcile()
    }
  }

  const isPending = (docId, flag) => pendingFlags.has(docId + '|' + flag)

  // Live popover cells derived from current `data` (activeDay/activeWeek are
  // keys). An optimistic flip mutates `data`, so the open modal re-renders
  // with fresh pill state — no close/reopen needed.
  const activeDayCell = activeDay ? (data?.days?.find(d => d.date === activeDay) || null) : null
  const activeWeekCell = activeWeek
    ? (data?.weeks?.find(w => w.week_start === activeWeek)
        || (activeWeek === lastWeekStart ? { week_start: activeWeek, breakdown: [] } : null))
    : null

  const onPendingMark = (docId, flag, value) => flip(docId, flag, value)

  const onPendingGenerate = (item) => { setPaystub(null); setPendingItem(item) }
  const closeGenerateModal = () => {
    setPendingItem(null)
    setPaystub(null)
    // Refetch so any state changes (e.g. Doc Lifecycle Log row created
    // by an upsert during generation) are reflected.
    load(monthIso)
  }
  const handleGenerateConfirm = async () => {
    if (!pendingItem) return
    const isCp = pendingItem.missing[0] === 'CP Done'
    const url = isCp
      ? '/api/tools/generate-cp-for-doc'
      : '/api/tools/generate-pl-for-doc'
    const res = await fetch(url, {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({
        doc_id: pendingItem.doc_id,
        // Optional paystub auto-fill — CP only, when one was uploaded.
        ...(isCp && paystub ? { paystub: { employees: paystub } } : {}),
      }),
    })
    const body = await res.json().catch(() => ({}))
    if (!res.ok || body.error) {
      throw new Error(body.error || `Request failed (HTTP ${res.status})`)
    }
    return { message: body.message, files: body.files || [], warnings: body.warnings || [] }
  }

  return (
    <div className="space-y-6">
      {/* Section — WO Docs: completed WOs whose CFR/Invoice isn't done+sent.
          Same wos array + write endpoints as the WO Tracker → 100% in sync. */}
      <div className="space-y-3">
        <h3 className="text-base font-black text-navy border-b border-slate-200 pb-1.5">WO Docs</h3>
        <WODocsQueue
          wos={wos}
          qbConnected={qbConnected}
          onDocsChange={onDocsChange}
          onInvoiced={onInvoiced}
        />
      </div>

      {/* Section — calendar-tracked docs, with its header right above the
          month selector. */}
      <div className="space-y-3">
        <h3 className="text-base font-black text-navy border-b border-slate-200 pb-1.5">Daily, Weekly, Monthly Docs</h3>
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
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <div className="space-y-3">
          <p className="section-label">Production Logs + Sign-Ins</p>
          <DayCalendar monthIso={monthIso} days={data?.days} loading={loading} onCellClick={setActiveDay} />
          <PendingList kind="day" pending={data?.pending} loading={loading} onMark={onPendingMark} onGenerate={onPendingGenerate} />
        </div>
        <div className="space-y-3">
          <p className="section-label">Certified Payroll + Month-End Docs</p>
          <WeekCalendar monthIso={monthIso} weeks={data?.weeks} monthEnd={data?.month_end} lastWeekStart={lastWeekStart} loading={loading} onCellClick={setActiveWeek} />
          <PendingList kind="week" pending={data?.pending} loading={loading} onMark={onPendingMark} onGenerate={onPendingGenerate} />
        </div>
      </div>

      {activeDayCell && (
        <DayCellPopover
          cell={activeDayCell}
          onClose={() => setActiveDay(null)}
          onFlip={flip}
          isPending={isPending}
        />
      )}
      {activeWeekCell && (
        <WeekCellPopover
          cell={activeWeekCell}
          monthEnd={data?.month_end}
          showMonthEnd={activeWeek === lastWeekStart}
          onClose={() => setActiveWeek(null)}
          onFlip={flip}
          isPending={isPending}
        />
      )}

      {flipError && (
        <div className="fixed bottom-4 left-1/2 -translate-x-1/2 z-50 bg-red-600 text-white text-sm font-semibold px-4 py-2 rounded-xl shadow-lg">
          ⚠ {flipError}
        </div>
      )}

      <GenerateDocModal
        open={!!pendingItem}
        title={pendingItem?.missing?.[0] === 'CP Done'
          ? 'Generate Certified Payroll'
          : 'Generate Production Log'}
        description={pendingItem ? (
          <span>
            <span className="block font-semibold text-slate-700">
              {pendingItem.missing?.[0] === 'CP Done'
                ? `Week of ${fmtShortDate(pendingItem.anchor)}`
                : fmtShortDate(pendingItem.anchor)}
            </span>
            <span className="block text-slate-500">
              {pendingItem.contract_num
                ? `${pendingItem.contractor} · ${pendingItem.contract_num} · ${pendingItem.borough}`
                : pendingItem.contractor}
            </span>
            <span className="block mt-2 text-[11px] text-slate-400">
              {pendingItem.missing?.[0] === 'PL Done'
                ? 'Covers all of this contractor’s contracts for the day. Review and sign the filled PDF in Approvals — Done flips automatically when archived.'
                : 'Generates the JSON template. Review and sign the filled PDF in the Approvals tab to mark Done.'}
            </span>
          </span>
        ) : null}
        onConfirm={handleGenerateConfirm}
        onClose={closeGenerateModal}
        idleExtra={pendingItem?.missing?.[0] === 'CP Done'
          ? <PaystubUpload onParsed={setPaystub} />
          : null}
      />
    </div>
  )
}
