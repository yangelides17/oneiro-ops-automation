import { useEffect, useState } from 'react'
import StatusBadge from './StatusBadge'

const STATUSES = ['Received', 'Dispatched', 'In Progress', 'Completed', 'Returned']

/**
 * StatusPickerModal — modal for the Dashboard WO kebab → "Change Status…"
 * action. Admin can flip a WO to any of the five statuses (including
 * the new "Returned" value), bypassing the one-way state machine that
 * the Field Report submit uses.
 *
 * Props:
 *   wo          — { id, status, location, contractor, ... }
 *   onSaved     — () => void  (parent refreshes Dashboard data)
 *   onClose     — () => void
 */
export default function StatusPickerModal({ wo, onSaved, onClose }) {
  const [picked, setPicked]       = useState(wo?.status || 'Received')
  const [submitting, setSubmitting] = useState(false)
  const [error, setError]         = useState('')

  useEffect(() => {
    const handler = (e) => { if (e.key === 'Escape' && !submitting) onClose?.() }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [onClose, submitting])

  if (!wo) return null

  const save = async () => {
    if (picked === wo.status) { onClose?.(); return }
    setSubmitting(true)
    setError('')
    try {
      const res = await fetch(`/api/wo/${encodeURIComponent(wo.id)}/status`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body:   JSON.stringify({ status: picked }),
      })
      const body = await res.json().catch(() => ({}))
      if (!res.ok || body.error) throw new Error(body.error || `HTTP ${res.status}`)
      onSaved?.()
      onClose?.()
    } catch (e) {
      setError(e.message || 'Failed to update status')
    } finally {
      setSubmitting(false)
    }
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4"
      style={{ backgroundColor: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(2px)' }}
      onClick={() => !submitting && onClose?.()}
    >
      <div
        className="bg-white rounded-2xl shadow-2xl w-full max-w-sm p-6 space-y-4"
        onClick={e => e.stopPropagation()}
      >
        <div className="text-center space-y-1.5">
          <h2 className="text-lg font-black text-navy">Change Status</h2>
          <p className="text-slate-500 text-sm leading-relaxed">
            <span className="font-mono font-bold text-navy">{wo.id}</span>
            {wo.location && <> — {wo.location}</>}
          </p>
        </div>

        <div className="space-y-1.5">
          {STATUSES.map(s => (
            <button
              key={s}
              type="button"
              disabled={submitting}
              onClick={() => setPicked(s)}
              className={`w-full text-left px-3 py-2.5 rounded-lg border transition-colors
                flex items-center justify-between
                ${picked === s
                  ? 'bg-navy/5 border-navy/40'
                  : 'bg-white border-slate-200 hover:border-navy/30'
                }`}
            >
              <span className="flex items-center gap-2">
                <span className={`w-4 h-4 rounded-full border-2 flex items-center justify-center
                                  ${picked === s ? 'border-navy' : 'border-slate-300'}`}>
                  {picked === s && <span className="w-2 h-2 rounded-full bg-navy" />}
                </span>
                <span className="text-sm font-semibold text-slate-700">{s}</span>
              </span>
              <StatusBadge status={s} />
            </button>
          ))}
        </div>

        {error && (
          <p className="text-xs text-red-600 font-semibold text-center">{error}</p>
        )}

        <div className="flex flex-col gap-2 pt-1">
          <button
            type="button"
            disabled={submitting || picked === wo.status}
            onClick={save}
            className="w-full py-3 rounded-xl font-bold text-sm bg-navy text-white
                       hover:opacity-90 active:opacity-80 transition-all
                       disabled:opacity-40 disabled:cursor-not-allowed"
          >
            {submitting ? 'Saving…' : 'Save'}
          </button>
          <button
            type="button"
            disabled={submitting}
            onClick={onClose}
            className="w-full py-3 rounded-xl font-bold text-sm bg-slate-100
                       text-slate-600 hover:bg-slate-200 transition-all
                       disabled:opacity-40"
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  )
}
