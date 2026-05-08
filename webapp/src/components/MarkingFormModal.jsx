import { useEffect, useMemo, useRef, useState } from 'react'
import {
  MARKING_CATEGORIES, UNIT_OPTIONS,
  unitForCategory, unitIsLocked, pickLayout,
} from '../lib/markingCategories'
import { parseQty } from '../lib/parseQty'
import { validateQty } from '../lib/qtyValidation'
import ConfirmModal from './ConfirmModal'

const DIRECTIONS = ['', 'N', 'E', 'S', 'W']

/**
 * MarkingFormModal — shared create/edit dialog for Marking Items.
 *
 * Visible inputs adapt to the picked Marking Type via pickLayout():
 *   mma     — Color/Material + Qty + Unit (locked SF) + Description + Notes
 *   grid    — Intersection + Direction + Qty + Unit (locked) + Description + Notes
 *   default — Qty + Unit + Description + Notes
 *
 * Marking Type is a custom searchable dropdown (no native datalist) so
 * mobile crews don't get word-prediction over the keyboard, and so users
 * can't type their own custom names that route incorrectly downstream.
 *
 * Intersection is a combobox of the WO's intersections (passed in via
 * `wo_intersections`) + derived between-pairs (`wo_betweens`) + an
 * "Other" option that flips to free text.
 *
 * Qty accepts arithmetic (`15*10`, `5+5+5`) via parseQty + runs the
 * same out-of-range check (validateQty) the inline grid uses on
 * HVX Crosswalk / Stop Line / Stop Msg.
 *
 * Props:
 *   mode             — 'add' | 'edit'
 *   item             — the existing item object (edit mode) or null (add mode)
 *   woId             — required in add mode
 *   workType         — 'MMA' | 'Thermo' | '' — drives layout fallback
 *   wo_intersections — string[] of distinct intersection names (in WO order)
 *   wo_betweens      — string[] of derived "X – Y" between-pairs
 *   onClose          — close without saving
 *   onSaved          — (updatedItem) => void   — called after a successful API call
 */
export default function MarkingFormModal({
  mode, item, woId, workType,
  wo_intersections = [], wo_betweens = [],
  onClose, onSaved,
}) {
  const isEdit          = mode === 'edit'
  const isAutoPopulated = isEdit && item?.added_by === 'Scanner'

  const [form, setForm] = useState(() => ({
    category:       item?.category       ?? '',
    intersection:   item?.intersection   ?? '',
    direction:      item?.direction      ?? '',
    description:    item?.description    ?? '',
    quantity:       item?.quantity != null ? String(item.quantity) : '',
    unit:           item?.unit           || 'SF',
    color_material: item?.color_material ?? '',
    notes:          item?.notes          ?? '',
  }))
  const [saving, setSaving] = useState(false)
  const [error,  setError]  = useState('')
  // Pending confirmation when validateQty flags an out-of-range value —
  // shape: { message, parsedStr } or null.
  const [qtyConfirm, setQtyConfirm] = useState(null)

  const layout = useMemo(
    () => pickLayout({ category: form.category, work_type: workType }),
    [form.category, workType]
  )

  // Close on Escape (only when nothing is mid-flight).
  useEffect(() => {
    const handler = (e) => {
      if (e.key !== 'Escape') return
      if (saving || qtyConfirm) return
      onClose?.()
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [saving, qtyConfirm, onClose])

  // When category changes, auto-update the unit from the category map
  // (variable categories like "Others" keep the user's last pick).
  const setField = (k, v) => setForm(f => {
    const next = { ...f, [k]: v }
    if (k === 'category') {
      const derived = unitForCategory(v)
      if (derived) next.unit = derived
      else if (!next.unit) next.unit = 'EA'
    }
    return next
  })

  // Build the final payload + actually POST/PATCH it. Called either
  // directly from handleConfirm (when no Qty warning) or from the
  // ConfirmModal's confirm callback (when admin OK'd a flagged Qty).
  async function persist(qtyParsedStr) {
    setError('')
    const category = form.category.trim()
    if (!category) { setError('Marking Type is required.'); return }

    const finalUnit = unitForCategory(category) || form.unit || 'EA'
    const qtyNum = (qtyParsedStr === '' || qtyParsedStr == null)
      ? null
      : parseFloat(qtyParsedStr)

    const payload = {
      category,
      intersection:   form.intersection.trim(),
      direction:      form.direction,
      description:    form.description.trim(),
      quantity:       qtyNum,
      unit:           finalUnit,
      color_material: form.color_material.trim(),
      notes:          form.notes.trim(),
    }

    setSaving(true)
    try {
      let res
      if (isEdit) {
        res = await fetch(`/api/marking-items/${encodeURIComponent(item.item_id)}`, {
          method:  'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body:    JSON.stringify(payload),
        })
      } else {
        res = await fetch('/api/marking-items', {
          method:  'POST',
          headers: { 'Content-Type': 'application/json' },
          body:    JSON.stringify({ ...payload, wo_id: woId, work_type: workType || '' }),
        })
      }
      const data = await res.json()
      if (data.error) throw new Error(data.error)
      if (!data.item) throw new Error('Server did not return an item')
      onSaved?.(data.item)
    } catch (e) {
      setError(e.message || 'Save failed')
      setSaving(false)
    }
  }

  // Save click — parses Qty arithmetic, runs the same out-of-range
  // validation the inline grid uses, then either prompts for confirm
  // (out of range) or persists immediately.
  function handleConfirm() {
    setError('')
    if (!form.category.trim()) {
      setError('Marking Type is required.')
      return
    }
    const parsedStr = parseQty(form.quantity)
    if (parsedStr !== form.quantity) {
      // Reflect the resolved number in the input so the modal message
      // and the eventually-saved value align with what the user sees.
      setField('quantity', parsedStr)
    }
    const check = validateQty(form.category, parsedStr)
    if (!check.ok) {
      setQtyConfirm({ message: check.message, parsedStr })
      return
    }
    persist(parsedStr)
  }

  return (
    <>
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4"
      style={{ backgroundColor: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(2px)' }}
      onClick={() => { if (!saving && !qtyConfirm) onClose?.() }}
    >
      <div
        className="bg-white rounded-2xl shadow-2xl w-full max-w-md p-5 space-y-4 max-h-[92vh] overflow-y-auto"
        onClick={e => e.stopPropagation()}
      >
        <div>
          <h2 className="text-lg font-black text-navy">
            {isEdit ? 'Edit Marking Item' : 'Add Marking Item'}
          </h2>
          {isEdit && (
            <p className="text-[11px] font-mono text-slate-400 mt-0.5">{item?.item_id}</p>
          )}
        </div>

        {isAutoPopulated && (
          <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 text-xs text-amber-800 leading-relaxed">
            <span className="font-bold">⚠ Auto-populated item.</span>{' '}
            This row was extracted from the WO scan. Changes here will overwrite the scanned values.
          </div>
        )}

        {error && (
          <div className="bg-red-50 border border-red-200 rounded-xl p-3 text-xs text-red-700">
            {error}
          </div>
        )}

        {/* Form */}
        <div className="space-y-2.5">
          <Field label="Marking Type" required>
            <CategorySelect
              value={form.category}
              onChange={(v) => setField('category', v)}
            />
          </Field>

          {layout === 'grid' && (
            <div className="grid grid-cols-[1fr_80px] gap-2">
              <Field label="Intersection">
                <IntersectionSelect
                  value={form.intersection}
                  onChange={(v) => setField('intersection', v)}
                  intersections={wo_intersections}
                  betweens={wo_betweens}
                />
              </Field>
              <Field label="Direction">
                <select
                  value={form.direction}
                  onChange={e => setField('direction', e.target.value)}
                  className="field-input"
                >
                  {DIRECTIONS.map(d => <option key={d} value={d}>{d || '—'}</option>)}
                </select>
              </Field>
            </div>
          )}

          {layout === 'mma' && (
            <Field label="Color / Material">
              <input
                type="text"
                value={form.color_material}
                onChange={e => setField('color_material', e.target.value)}
                placeholder="e.g. White Thermo"
                className="field-input"
              />
            </Field>
          )}

          <div className="grid grid-cols-[1fr_80px] gap-2">
            <Field label="Quantity">
              <input
                type="text"
                inputMode="text"
                value={form.quantity}
                onChange={e => {
                  // Same character allowlist the inline grid uses so paste
                  // / stray taps can't land letters / units in the field.
                  const v = (e.target.value.match(/[\d.+*xX\s]/g) || []).join('')
                  setField('quantity', v)
                }}
                placeholder="0  (or 15*10, 5+5+5)"
                className="field-input"
              />
            </Field>
            <Field label="Unit">
              {unitIsLocked(form.category) ? (
                <div className="field-input bg-slate-50 text-slate-500
                                flex items-center justify-center font-semibold
                                cursor-not-allowed">
                  {unitForCategory(form.category)}
                </div>
              ) : (
                <select
                  value={form.unit}
                  onChange={e => setField('unit', e.target.value)}
                  className="field-input"
                >
                  {UNIT_OPTIONS.map(u => <option key={u}>{u}</option>)}
                </select>
              )}
            </Field>
          </div>

          <Field label="Description">
            <input
              type="text"
              value={form.description}
              onChange={e => setField('description', e.target.value)}
              placeholder="e.g. RECAP FROM HAMILTON PL TO 2ND AV"
              className="field-input"
            />
          </Field>

          <Field label="Notes">
            <textarea
              value={form.notes}
              onChange={e => setField('notes', e.target.value)}
              rows={2}
              placeholder="Optional notes for the admin…"
              className="field-input resize-none"
            />
          </Field>
        </div>

        {/* Buttons */}
        <div className="flex flex-col gap-2 pt-1">
          <button
            type="button"
            onClick={handleConfirm}
            disabled={saving}
            className="w-full py-3 rounded-xl font-bold text-sm
                       bg-navy text-white hover:opacity-90 active:opacity-80
                       disabled:opacity-60 disabled:cursor-not-allowed transition-all"
          >
            {saving
              ? 'Saving…'
              : isEdit ? 'Confirm changes' : 'Confirm & add'}
          </button>
          <button
            type="button"
            onClick={onClose}
            disabled={saving}
            className="w-full py-3 rounded-xl font-bold text-sm bg-slate-100
                       text-slate-600 hover:bg-slate-200 disabled:opacity-60
                       disabled:cursor-not-allowed transition-all"
          >
            Cancel
          </button>
        </div>
      </div>
    </div>

    {qtyConfirm && (
      <ConfirmModal
        title="Quantity outside typical range"
        message={qtyConfirm.message}
        confirmLabel="Yes, keep it"
        cancelLabel="Edit value"
        onConfirm={() => {
          const { parsedStr } = qtyConfirm
          setQtyConfirm(null)
          persist(parsedStr)
        }}
        onCancel={() => setQtyConfirm(null)}
      />
    )}
    </>
  )
}

function Field({ label, required, children }) {
  return (
    <div className="space-y-1">
      <label className="field-label">
        {label}{required && <span className="text-red-400 ml-0.5">*</span>}
      </label>
      {children}
    </div>
  )
}

// ── CategorySelect ────────────────────────────────────────────────
// Custom searchable dropdown for Marking Type. Replaces the native
// <input list="datalist"> which on mobile shows word-prediction
// suggestions over the keyboard and lets users type custom names.
//
// No free-text on the trigger — picking from the list is the only
// way to set the category. Edit-mode preserves legacy values not in
// MARKING_CATEGORIES via a "Legacy: <value>" sentinel option.
function CategorySelect({ value, onChange }) {
  const [open, setOpen]     = useState(false)
  const [query, setQuery]   = useState('')
  const wrapRef = useRef(null)
  const searchRef = useRef(null)

  useEffect(() => {
    if (!open) return
    setQuery('')
    setTimeout(() => searchRef.current?.focus(), 0)
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

  const isLegacy = value && MARKING_CATEGORIES.indexOf(value) === -1
  const opts = useMemo(() => {
    const q = query.trim().toLowerCase()
    const list = isLegacy ? [`Legacy: ${value}`, ...MARKING_CATEGORIES] : MARKING_CATEGORIES
    if (!q) return list
    return list.filter(c => c.toLowerCase().indexOf(q) !== -1)
  }, [query, isLegacy, value])

  const pick = (c) => {
    if (c.startsWith('Legacy: ')) {
      // Keep the legacy value verbatim.
      onChange(c.slice('Legacy: '.length))
    } else {
      onChange(c)
    }
    setOpen(false)
  }

  return (
    <div ref={wrapRef} className="relative">
      <button
        type="button"
        onClick={() => setOpen(o => !o)}
        className="field-input w-full text-left flex items-center justify-between"
      >
        <span className={value ? 'text-slate-800' : 'text-slate-400'}>
          {value || 'Select a marking type'}
        </span>
        <span className="text-slate-400 text-xs ml-2">▾</span>
      </button>
      {open && (
        <div className="absolute z-30 mt-1 left-0 right-0 bg-white rounded-xl shadow-lg border border-slate-200 overflow-hidden">
          <div className="p-2 border-b border-slate-100">
            <input
              ref={searchRef}
              type="text"
              value={query}
              onChange={e => setQuery(e.target.value)}
              placeholder="Search marking types…"
              className="field-input text-base sm:text-sm"
            />
          </div>
          <div className="max-h-[260px] overflow-y-auto">
            {opts.length === 0 && (
              <p className="px-3 py-2 text-xs text-slate-400 italic">No matches</p>
            )}
            {opts.map((c) => {
              const isSelected = (isLegacy && c === `Legacy: ${value}`) || c === value
              return (
                <button
                  key={c}
                  type="button"
                  onClick={() => pick(c)}
                  className={`w-full text-left px-3 py-2 text-sm
                              hover:bg-slate-50
                              ${isSelected ? 'bg-navy/5 font-semibold text-navy' : 'text-slate-700'}`}
                >
                  {c}
                </button>
              )
            })}
          </div>
        </div>
      )}
    </div>
  )
}

// ── IntersectionSelect ────────────────────────────────────────────
// Combobox: search + scrollable list of the WO's intersections, then
// derived between-pairs ("X – Y"), then "Other" → free-text.
function IntersectionSelect({ value, onChange, intersections = [], betweens = [] }) {
  const [open, setOpen]   = useState(false)
  const [query, setQuery] = useState('')
  const [customMode, setCustomMode] = useState(false)
  const wrapRef = useRef(null)
  const searchRef = useRef(null)
  const customRef = useRef(null)

  useEffect(() => {
    if (!open) return
    setQuery('')
    setTimeout(() => {
      if (customMode) customRef.current?.focus()
      else searchRef.current?.focus()
    }, 0)
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
  }, [open, customMode])

  const filterMatch = (s) => !query.trim() || s.toLowerCase().indexOf(query.trim().toLowerCase()) !== -1
  const filteredIntersections = intersections.filter(filterMatch)
  const filteredBetweens      = betweens.filter(filterMatch)

  const pick = (v) => {
    onChange(v)
    setOpen(false)
    setCustomMode(false)
  }

  return (
    <div ref={wrapRef} className="relative">
      <button
        type="button"
        onClick={() => { setOpen(o => !o); setCustomMode(false) }}
        className="field-input w-full text-left flex items-center justify-between"
      >
        <span className={value ? 'text-slate-800' : 'text-slate-400'}>
          {value || 'Select an intersection'}
        </span>
        <span className="text-slate-400 text-xs ml-2">▾</span>
      </button>
      {open && (
        <div className="absolute z-30 mt-1 left-0 right-0 bg-white rounded-xl shadow-lg border border-slate-200 overflow-hidden">
          {customMode ? (
            <div className="p-2 space-y-2">
              <input
                ref={customRef}
                type="text"
                defaultValue={value || ''}
                placeholder="e.g. 5 AV"
                className="field-input text-base sm:text-sm"
                onKeyDown={e => {
                  if (e.key === 'Enter') {
                    pick(e.target.value.trim())
                  }
                }}
              />
              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={() => pick(customRef.current?.value.trim() || '')}
                  className="flex-1 py-2 rounded-lg text-xs font-bold bg-navy text-white hover:opacity-90"
                >
                  Save
                </button>
                <button
                  type="button"
                  onClick={() => setCustomMode(false)}
                  className="flex-1 py-2 rounded-lg text-xs font-bold bg-slate-100 text-slate-600 hover:bg-slate-200"
                >
                  Back
                </button>
              </div>
            </div>
          ) : (
            <>
              <div className="p-2 border-b border-slate-100">
                <input
                  ref={searchRef}
                  type="text"
                  value={query}
                  onChange={e => setQuery(e.target.value)}
                  placeholder="Search intersections…"
                  className="field-input text-base sm:text-sm"
                />
              </div>
              <div className="max-h-[260px] overflow-y-auto">
                {filteredIntersections.length > 0 && (
                  <>
                    <p className="px-3 pt-2 pb-1 text-[10px] font-extrabold uppercase tracking-widest text-slate-400">
                      Intersections
                    </p>
                    {filteredIntersections.map((c) => (
                      <button
                        key={'i-' + c}
                        type="button"
                        onClick={() => pick(c)}
                        className={`w-full text-left px-3 py-2 text-sm hover:bg-slate-50
                                    ${c === value ? 'bg-navy/5 font-semibold text-navy' : 'text-slate-700'}`}
                      >
                        {c}
                      </button>
                    ))}
                  </>
                )}
                {filteredBetweens.length > 0 && (
                  <>
                    <p className="px-3 pt-2 pb-1 text-[10px] font-extrabold uppercase tracking-widest text-slate-400">
                      Between
                    </p>
                    {filteredBetweens.map((c) => (
                      <button
                        key={'b-' + c}
                        type="button"
                        onClick={() => pick(c)}
                        className={`w-full text-left px-3 py-2 text-sm hover:bg-slate-50
                                    ${c === value ? 'bg-navy/5 font-semibold text-navy' : 'text-slate-700'}`}
                      >
                        {c}
                      </button>
                    ))}
                  </>
                )}
                {filteredIntersections.length === 0 && filteredBetweens.length === 0 && (
                  <p className="px-3 py-2 text-xs text-slate-400 italic">
                    No matches — pick "Other" to enter a custom value.
                  </p>
                )}
                <div className="border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setCustomMode(true)}
                    className="w-full text-left px-3 py-2 text-sm font-semibold text-navy hover:bg-slate-50"
                  >
                    Other (type a custom value)…
                  </button>
                </div>
              </div>
            </>
          )}
        </div>
      )}
    </div>
  )
}
