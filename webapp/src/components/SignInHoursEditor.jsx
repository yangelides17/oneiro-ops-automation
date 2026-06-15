import { useEffect, useMemo, useState } from 'react'
import { CLASSIFICATIONS, calcHours, splitStOt, to24h } from '../lib/signinShared'

// Approvals-page editor for a submitted sign-in's recorded hours. Lets
// an admin correct Time In / Time Out / Classification before approving
// so what's in the system matches the signed sheet. Writes back to Daily
// Sign-In Data only (the payroll source of truth) — the signed PDF is
// never touched. Hours + Overtime recompute automatically (weekend/8-hr
// rule) and the whole day's OT is re-split server-side on save.
//
// Props:
//   fileId, filename  — the pending sign-in being reviewed
//   onDirtyChange(bool) — lifts unsaved-edits state so the parent can
//                         warn before an approve wipes the edits.

const rowToEdit = (r) => ({
  row_index:      r.row_index,
  name:           r.name,
  classification: r.classification || CLASSIFICATIONS[0],
  timeIn:         to24h(r.time_in),
  timeOut:        to24h(r.time_out),
  crew_chief:     r.crew_chief || '',
})

export default function SignInHoursEditor({ fileId, filename, onDirtyChange }) {
  const [loading, setLoading]   = useState(true)
  const [error,   setError]     = useState('')
  const [meta,    setMeta]      = useState(null)
  const [otherHours, setOtherHours] = useState({})
  const [edits,   setEdits]     = useState([])      // live, editable
  const [original, setOriginal] = useState([])      // snapshot for dirty-check
  const [saving,  setSaving]    = useState(false)
  const [saveError, setSaveError] = useState('')
  const [savedNote, setSavedNote] = useState('')

  // ── Load rows whenever the selected file changes ─────────────
  useEffect(() => {
    let cancelled = false
    setLoading(true); setError(''); setSaveError(''); setSavedNote('')
    const qs = new URLSearchParams({ filename: filename || '' }).toString()
    fetch(`/api/approvals/${encodeURIComponent(fileId)}/signin-rows?${qs}`)
      .then(r => r.json())
      .then(json => {
        if (cancelled) return
        if (json.error) { setError(json.error); setEdits([]); setOriginal([]); return }
        const rows = Array.isArray(json.rows) ? json.rows : []
        setMeta(json.meta || null)
        setOtherHours(json.other_hours || {})
        setEdits(rows.map(rowToEdit))
        setOriginal(rows.map(rowToEdit))
      })
      .catch(err => { if (!cancelled) setError(err.message || 'Failed to load hours') })
      .finally(() => { if (!cancelled) setLoading(false) })
    return () => { cancelled = true }
  }, [fileId, filename])

  const dateIso   = meta?.date || ''
  const ambiguous = !!meta?.ambiguous

  // ── Dirty tracking ───────────────────────────────────────────
  const isDirty = useMemo(() => {
    if (edits.length !== original.length) return false
    return edits.some((e, i) => {
      const o = original[i]
      return !o || e.classification !== o.classification ||
        e.timeIn !== o.timeIn || e.timeOut !== o.timeOut
    })
  }, [edits, original])

  useEffect(() => { onDirtyChange?.(isDirty) }, [isDirty, onDirtyChange])
  // Clear the dirty flag in the parent when this editor unmounts (e.g.
  // the doc is approved and leaves the list).
  useEffect(() => () => onDirtyChange?.(false), [onDirtyChange])

  const setRow = (idx, patch) =>
    setEdits(prev => prev.map((e, i) => (i === idx ? { ...e, ...patch } : e)))

  // ── Per-row live hours/OT (display only; server authoritative) ─
  const rowCalc = (e) => calcHours(e.timeIn, e.timeOut, dateIso)

  // ── Shift totals for the day (this sheet live + other sheets) ─
  const totals = useMemo(() => {
    const thisSheet = {}
    edits.forEach(e => {
      const name = (e.name || '').trim()
      if (!name) return
      const h = parseFloat(rowCalc(e).hours) || 0
      thisSheet[name] = (thisSheet[name] || 0) + h
    })
    const names = new Set([...Object.keys(thisSheet), ...Object.keys(otherHours)])
    const out = []
    names.forEach(name => {
      const mine  = thisSheet[name] || 0
      const other = Number(otherHours[name] || 0)
      const total = mine + other
      out.push({ name, thisSheet: mine, otherSheets: other, total, ...splitStOt(total, dateIso) })
    })
    out.sort((a, b) => a.name.localeCompare(b.name))
    return out
  }, [edits, otherHours, dateIso])

  // ── Save ─────────────────────────────────────────────────────
  const handleSave = async () => {
    if (saving) return
    setSaveError(''); setSavedNote('')
    // Validate: every edited row needs both times.
    for (const e of edits) {
      if (!e.timeIn || !e.timeOut) {
        setSaveError(`${e.name || 'A crew member'} needs both Time In and Time Out`)
        return
      }
    }
    // Only send rows that actually changed.
    const changed = edits.filter((e, i) => {
      const o = original[i]
      return !o || e.classification !== o.classification ||
        e.timeIn !== o.timeIn || e.timeOut !== o.timeOut
    })
    if (!changed.length) return
    setSaving(true)
    try {
      const res = await fetch(`/api/approvals/${encodeURIComponent(fileId)}/save-signin-rows`, {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          filename,
          rows: changed.map(e => ({
            row_index:      e.row_index,
            classification: e.classification,
            time_in:        e.timeIn,
            time_out:       e.timeOut,
          })),
        }),
      })
      const json = await res.json()
      if (!res.ok || json.error) throw new Error(json.error || `HTTP ${res.status}`)
      const rows = Array.isArray(json.rows) ? json.rows : []
      setEdits(rows.map(rowToEdit))
      setOriginal(rows.map(rowToEdit))
      const n = json.other_sheet_updates || 0
      setSavedNote(n > 0
        ? `Saved. Also recomputed overtime on ${n} row(s) on other sheets for this day.`
        : 'Saved.')
    } catch (err) {
      setSaveError(err.message || 'Save failed')
    } finally {
      setSaving(false)
    }
  }

  // ── Render ───────────────────────────────────────────────────
  if (loading) {
    return (
      <div className="mb-3 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-500">
        Loading recorded hours…
      </div>
    )
  }
  if (error) {
    return (
      <div className="mb-3 rounded-lg border border-amber-200 bg-amber-50 px-3 py-2 text-xs text-amber-800">
        Couldn’t load hours for this sheet: {error}
      </div>
    )
  }
  if (ambiguous) {
    return (
      <div className="mb-3 rounded-lg border border-amber-200 bg-amber-50 px-3 py-2 text-xs text-amber-800">
        <span className="font-semibold">Multiple crews on this date</span> — this older sheet
        doesn’t identify which crew it belongs to, so hours can’t be safely edited here. Adjust
        them directly in the Daily Sign-In Data spreadsheet.
      </div>
    )
  }
  if (!edits.length) {
    return (
      <div className="mb-3 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-500">
        No recorded hours found in Daily Sign-In Data for this sheet.
      </div>
    )
  }

  return (
    <div className="mb-3 rounded-xl border border-slate-200 bg-white">
      <div className="flex items-center justify-between gap-2 px-3 py-2 border-b border-slate-100">
        <p className="text-xs font-bold uppercase tracking-wider text-slate-500">
          Recorded Hours {meta?.date && <span className="text-slate-400">· {meta.date}</span>}
        </p>
        {isDirty && <span className="text-[11px] font-semibold text-orange-600">Unsaved edits</span>}
      </div>

      {/* Editable rows */}
      <div className="divide-y divide-slate-100">
        {/* header — OT is intentionally NOT per-row here: a worker's OT
            depends on their whole-day total across sheets, so the
            accurate split lives in the Shift Totals strip below. */}
        <div className="hidden sm:grid grid-cols-12 gap-2 px-3 py-1.5 text-[10px] uppercase tracking-wider font-bold text-slate-400">
          <span className="col-span-4">Employee</span>
          <span className="col-span-2">Class</span>
          <span className="col-span-3">In</span>
          <span className="col-span-2">Out</span>
          <span className="col-span-1 text-right">Hrs</span>
        </div>
        {edits.map((e, i) => {
          const c = rowCalc(e)
          return (
            <div key={e.row_index} className="grid grid-cols-12 gap-2 px-3 py-2 items-center">
              <span className="col-span-12 sm:col-span-4 text-sm font-semibold text-navy truncate">
                {e.name}
              </span>
              <div className="col-span-4 sm:col-span-2">
                <select
                  value={e.classification}
                  onChange={ev => setRow(i, { classification: ev.target.value })}
                  className="field-input !py-1 text-sm">
                  {CLASSIFICATIONS.map(cl => <option key={cl}>{cl}</option>)}
                </select>
              </div>
              <div className="col-span-4 sm:col-span-3">
                <input
                  type="time"
                  value={e.timeIn}
                  onChange={ev => setRow(i, { timeIn: ev.target.value })}
                  className="field-input !py-1 text-sm" />
              </div>
              <div className="col-span-4 sm:col-span-2">
                <input
                  type="time"
                  value={e.timeOut}
                  onChange={ev => setRow(i, { timeOut: ev.target.value })}
                  className="field-input !py-1 text-sm" />
              </div>
              <span className="col-span-12 sm:col-span-1 text-right text-sm font-semibold text-slate-700">
                {c.hours || '—'}
              </span>
            </div>
          )
        })}
      </div>

      {/* Shift totals strip */}
      <div className="px-3 py-2 border-t border-slate-100 bg-slate-50/60">
        <p className="text-[10px] uppercase tracking-wider font-bold text-slate-400 mb-1">
          Shift Totals · all sign-ins this date
        </p>
        <div className="space-y-1">
          {totals.map(t => (
            <div key={t.name} className="grid grid-cols-12 gap-2 text-[12px] items-baseline">
              <span className="col-span-5 font-medium text-slate-700 truncate">{t.name}</span>
              <span className="col-span-2 text-right text-slate-500" title="this sheet">{t.thisSheet.toFixed(2)}</span>
              <span className="col-span-2 text-right text-slate-400" title="other sheets">{t.otherSheets.toFixed(2)}</span>
              <span className="col-span-2 text-right font-semibold text-navy" title="ST">{t.st.toFixed(2)}</span>
              <span className={`col-span-1 text-right font-bold ${t.ot > 0 ? 'text-orange-600' : 'text-slate-300'}`} title="OT">{t.ot.toFixed(2)}</span>
            </div>
          ))}
        </div>
      </div>

      {/* Save */}
      <div className="flex items-center justify-between gap-3 px-3 py-2 border-t border-slate-100">
        <div className="text-[11px] min-w-0">
          {saveError && <span className="text-red-600">{saveError}</span>}
          {!saveError && savedNote && <span className="text-emerald-700">{savedNote}</span>}
        </div>
        <button
          type="button"
          onClick={handleSave}
          disabled={!isDirty || saving}
          className="text-xs font-bold px-4 py-1.5 rounded-lg bg-navy text-white
                     hover:opacity-90 disabled:opacity-40 disabled:cursor-not-allowed flex-shrink-0">
          {saving ? 'Saving…' : 'Save hours'}
        </button>
      </div>
    </div>
  )
}
