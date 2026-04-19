import { useEffect, useState } from 'react'

// Canonical marking categories — kept in sync with setupMarkingItems()
// in Code.js. Non-strict dropdown so users can type anything.
const MARKING_CATEGORIES = [
  // WO Top Table
  'Double Yellow Line', 'Lane Lines', 'Gores', 'Messages', 'Arrows',
  'Solid Lines', 'Rail Road X/Diamond', 'Others',
  // Intersection Grid
  'HVX Crosswalk', 'Stop Msg', 'Stop Line',
  // Page 2 detailed lines
  '4" Line', '6" Line', '8" Line', '12" Line', '16" Line', '24" Line',
  // Page 2 messages
  'Only Msg', 'Bus Msg', 'Bump Msg', 'Custom Msg', '20 MPH Msg',
  // Page 2 railroad
  'Railroad (RR)', 'Railroad (X)',
  // Page 2 arrows
  'L/R Arrow', 'Straight Arrow', 'Combination Arrow',
  // Page 2 misc
  'Speed Hump Markings', 'Shark Teeth 12x18', 'Shark Teeth 24x36',
  // Page 2 bike lane
  'Bike Lane Arrow', 'Bike Lane Symbol', 'Bike Lane Green Bar',
  // MMA
  'Bike Lane', 'Pedestrian Space', 'Bus Lane', 'Ped Stop',
]

const UNITS      = ['SF', 'LF', 'EA']
const DIRECTIONS = ['', 'N', 'E', 'S', 'W']

/**
 * MarkingFormModal — shared create/edit dialog for Marking Items.
 *
 * Props:
 *   mode      — 'add' | 'edit'
 *   item      — the existing item object (edit mode) or null (add mode)
 *   woId      — required in add mode
 *   workType  — 'MMA' | 'Thermo' | '' — drives whether Color/Material is shown
 *   onClose   — close without saving
 *   onSaved   — (updatedItem) => void   — called after a successful API call
 */
export default function MarkingFormModal({ mode, item, woId, workType, onClose, onSaved }) {
  const isEdit          = mode === 'edit'
  const isAutoPopulated = isEdit && item?.added_by === 'Scanner'
  const isMMA           = String(workType || '').toLowerCase() === 'mma'

  const [form, setForm] = useState(() => ({
    category:       item?.category       ?? '',
    intersection:   item?.intersection   ?? '',
    direction:      item?.direction      ?? '',
    description:    item?.description    ?? '',
    quantity:       item?.quantity ?? '',
    unit:           item?.unit           || 'SF',
    color_material: item?.color_material ?? '',
    notes:          item?.notes          ?? '',
  }))
  const [saving, setSaving] = useState(false)
  const [error,  setError]  = useState('')

  // Close on Escape
  useEffect(() => {
    const handler = (e) => { if (e.key === 'Escape' && !saving) onClose?.() }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [saving, onClose])

  const setField = (k, v) => setForm(f => ({ ...f, [k]: v }))

  async function handleConfirm() {
    setError('')
    const category = form.category.trim()
    if (!category) { setError('Marking Type is required.'); return }

    const payload = {
      category,
      intersection:   form.intersection.trim(),
      direction:      form.direction,
      description:    form.description.trim(),
      quantity:       form.quantity === '' ? null : parseFloat(form.quantity),
      unit:           form.unit || 'SF',
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

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4"
      style={{ backgroundColor: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(2px)' }}
      onClick={() => { if (!saving) onClose?.() }}
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
            <input
              list="marking-categories"
              type="text"
              value={form.category}
              onChange={e => setField('category', e.target.value)}
              placeholder="Select or type a category"
              className="field-input"
            />
            <datalist id="marking-categories">
              {MARKING_CATEGORIES.map(c => <option key={c} value={c} />)}
            </datalist>
          </Field>

          <div className="grid grid-cols-[1fr_80px] gap-2">
            <Field label="Intersection">
              <input
                type="text"
                value={form.intersection}
                onChange={e => setField('intersection', e.target.value)}
                placeholder="e.g. 5 AV"
                className="field-input"
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

          <Field label="Description">
            <input
              type="text"
              value={form.description}
              onChange={e => setField('description', e.target.value)}
              placeholder="e.g. RECAP FROM HAMILTON PL TO 2ND AV"
              className="field-input"
            />
          </Field>

          <div className="grid grid-cols-[1fr_80px] gap-2">
            <Field label="Quantity">
              <input
                type="number" min="0" inputMode="numeric"
                value={form.quantity}
                onChange={e => setField('quantity', e.target.value)}
                placeholder="0"
                className="field-input"
              />
            </Field>
            <Field label="Unit">
              <select
                value={form.unit}
                onChange={e => setField('unit', e.target.value)}
                className="field-input"
              >
                {UNITS.map(u => <option key={u}>{u}</option>)}
              </select>
            </Field>
          </div>

          {isMMA && (
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
