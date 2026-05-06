import { useEffect, useRef, useState } from 'react'

/**
 * DocStatusChips — compact 5-chip cluster showing per-doc-type
 * lifecycle state on a WO row. Click anywhere on the cluster to
 * open an inline editor popover with toggle buttons for Done/Sent
 * across all 5 doc types.
 *
 * Props:
 *   woId     — WO id (string), passed back through onChange
 *   docs     — { cfr, production_log, signin, certified_payroll, invoice },
 *              each { done: bool, sent: bool }
 *   onChange — (woId, friendlyDocType, partial) => Promise<bool>
 *              partial: { done?: bool } | { sent?: bool }
 *              Returns true on success, false on failure (so the
 *              chip can revert its optimistic update).
 */

// Doc-type metadata keyed by the same keys used in the dashboard payload.
// Keep `friendly` in sync with the strings the Apps Script set_docs_sent
// action accepts. Slimmed to CFR + INV — those are the only doc types
// that genuinely fit per-WO storage. Time-anchored docs (PL/SI/CP)
// moved to the new Doc Lifecycle Log; their state lives on the Doc
// Status tab calendar instead.
const DOC_TYPES = [
  { key: 'cfr',     friendly: 'CFR',     label: 'Contractor Field Report', short: 'CFR' },
  { key: 'invoice', friendly: 'Invoice', label: 'Invoice',                 short: 'INV' },
]

// Lifecycle state → visual style. Three states only: not done, done
// (not sent), sent. Sent implies done so we don't surface a separate
// "sent but not done" state — that's invariant-violating data.
function chipClass(done, sent) {
  if (sent) return 'bg-green-100 text-green-800 border-green-200'
  if (done) return 'bg-amber-100 text-amber-800 border-amber-200'
  return 'bg-slate-100 text-slate-500 border-slate-200'
}
function chipTooltip(label, done, sent) {
  return `${label} — Done: ${done ? 'Yes' : 'No'}, Sent: ${sent ? 'Yes' : 'No'}`
}

export default function DocStatusChips({ woId, docs, onChange }) {
  const [open, setOpen] = useState(false)
  // Local optimistic copy of the docs object — updated immediately on
  // toggle, revert if onChange returns false.
  const [local, setLocal] = useState(docs || {})
  // Per-toggle "in flight" guard so rapid double-clicks don't fire
  // duplicate fetches. Keyed by `${docKey}|${flag}`.
  const [pending, setPending] = useState({})
  const wrapRef = useRef(null)

  // Sync local state when parent prop changes (e.g. dashboard refetch).
  useEffect(() => { setLocal(docs || {}) }, [docs])

  // Outside-click + Escape to close (matches RowKebab pattern).
  useEffect(() => {
    if (!open) return
    const handler = (e) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) setOpen(false)
    }
    const esc = (e) => { if (e.key === 'Escape') setOpen(false) }
    document.addEventListener('mousedown', handler)
    document.addEventListener('keydown', esc)
    return () => {
      document.removeEventListener('mousedown', handler)
      document.removeEventListener('keydown', esc)
    }
  }, [open])

  // Toggle one flag for one doc type. Optimistic — flips local state
  // immediately, fires the network call, reverts on failure.
  async function toggle(docKey, flag) {
    const docMeta = DOC_TYPES.find(d => d.key === docKey)
    if (!docMeta) return
    const cur = local[docKey] || { done: false, sent: false }
    // Sent requires Done. The button is disabled in that case but
    // belt-and-suspenders here too.
    if (flag === 'sent' && !cur.done) return
    const pendKey = `${docKey}|${flag}`
    if (pending[pendKey]) return
    const next = { ...cur, [flag]: !cur[flag] }
    // Turning off Done while Sent is on — also clear Sent. Otherwise
    // we'd leave the row in an invariant-violating state.
    if (flag === 'done' && !next.done && cur.sent) next.sent = false

    setLocal(prev => ({ ...prev, [docKey]: next }))
    setPending(prev => ({ ...prev, [pendKey]: true }))
    try {
      const partial = { done: next.done, sent: next.sent }
      const ok = await onChange(woId, docMeta.friendly, partial)
      if (!ok) {
        // Revert
        setLocal(prev => ({ ...prev, [docKey]: cur }))
      }
    } finally {
      setPending(prev => {
        const copy = { ...prev }
        delete copy[pendKey]
        return copy
      })
    }
  }

  return (
    <div ref={wrapRef} className="relative inline-block">
      {/* Chip cluster — click to open editor */}
      <button
        type="button"
        onClick={(e) => { e.stopPropagation(); setOpen(o => !o) }}
        className="flex items-center gap-1 px-1 py-0.5 rounded-md
                   hover:bg-slate-50 transition-colors"
        aria-label="Edit document statuses"
      >
        {DOC_TYPES.map(d => {
          const s = local[d.key] || { done: false, sent: false }
          return (
            <span
              key={d.key}
              title={chipTooltip(d.label, s.done, s.sent)}
              className={`inline-flex items-center justify-center
                          text-[9px] font-extrabold tracking-wider
                          px-1.5 h-[18px] min-w-[24px] rounded
                          border ${chipClass(s.done, s.sent)}`}
            >
              {d.short}
            </span>
          )
        })}
      </button>

      {/* Editor popover */}
      {open && (
        <div
          className="absolute right-0 mt-1 z-30 min-w-[280px]
                     bg-white rounded-lg shadow-lg border border-slate-200
                     py-2 text-sm"
          onClick={e => e.stopPropagation()}
        >
          <div className="px-3 pb-2 mb-1 border-b border-slate-100 flex items-center justify-between">
            <span className="text-[11px] font-bold uppercase tracking-widest text-slate-500">
              {woId} · Documents
            </span>
            <button
              onClick={() => setOpen(false)}
              className="text-slate-400 hover:text-slate-600 text-base leading-none"
              aria-label="Close"
            >×</button>
          </div>
          <div className="px-3 py-1 space-y-2">
            {DOC_TYPES.map(d => {
              const s = local[d.key] || { done: false, sent: false }
              const sentDisabled = !s.done
              const donePending = pending[`${d.key}|done`]
              const sentPending = pending[`${d.key}|sent`]
              return (
                <div key={d.key} className="flex items-center justify-between gap-3">
                  <span className="text-xs text-slate-700 font-medium truncate">
                    {d.label}
                  </span>
                  <div className="flex gap-1.5 flex-shrink-0">
                    <ToggleButton
                      label="Done"
                      on={s.done}
                      pending={donePending}
                      onClick={() => toggle(d.key, 'done')}
                    />
                    <ToggleButton
                      label="Sent"
                      on={s.sent}
                      pending={sentPending}
                      disabled={sentDisabled}
                      title={sentDisabled ? 'Mark Done first' : ''}
                      onClick={() => toggle(d.key, 'sent')}
                    />
                  </div>
                </div>
              )
            })}
          </div>
        </div>
      )}
    </div>
  )
}

function ToggleButton({ label, on, pending, disabled, title, onClick }) {
  const base = 'text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded border transition-all'
  let style
  if (disabled) {
    style = 'bg-slate-50 text-slate-300 border-slate-100 cursor-not-allowed'
  } else if (on) {
    style = label === 'Sent'
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
      title={title}
      className={`${base} ${style} ${pending ? 'opacity-60' : ''}`}
    >
      {pending ? '…' : label}
    </button>
  )
}
