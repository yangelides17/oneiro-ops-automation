import { useEffect, useState } from 'react'

/**
 * DeleteWOModal — confirmation modal for the Dashboard WO kebab →
 * "Delete WO…" action. Hard-deletes the Tracker row + all Marking
 * Items rows keyed to the WO. Work Day Log entries, Sign-In Data,
 * Doc Lifecycle Log, and the Drive archive folder are preserved
 * (audit trail).
 *
 * Props:
 *   wo        — { id, status, location, contractor, ... }
 *   onDeleted — (result) => void  (result includes counts)
 *   onClose   — () => void
 */
export default function DeleteWOModal({ wo, onDeleted, onClose }) {
  const [submitting, setSubmitting] = useState(false)
  const [error, setError]           = useState('')

  useEffect(() => {
    const handler = (e) => { if (e.key === 'Escape' && !submitting) onClose?.() }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [onClose, submitting])

  if (!wo) return null

  // If the WO has been worked (Completed / In Progress) surface the
  // extra warning so the admin knows time-card / payroll records will
  // remain even though the WO row is gone.
  const wasWorked = wo.status &&
    ['completed', 'in progress'].indexOf(String(wo.status).toLowerCase()) !== -1

  const doDelete = async () => {
    setSubmitting(true)
    setError('')
    try {
      const res = await fetch(`/api/wo/${encodeURIComponent(wo.id)}`, { method: 'DELETE' })
      const body = await res.json().catch(() => ({}))
      if (!res.ok || body.error) throw new Error(body.error || `HTTP ${res.status}`)
      onDeleted?.(body)
      onClose?.()
    } catch (e) {
      setError(e.message || 'Failed to delete WO')
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
        className="bg-white rounded-2xl shadow-2xl w-full max-w-md p-6 space-y-4"
        onClick={e => e.stopPropagation()}
      >
        <div className="text-3xl text-center">⚠️</div>

        <div className="text-center space-y-1.5">
          <h2 className="text-lg font-black text-navy">Delete this Work Order?</h2>
          <p className="text-slate-600 text-sm">
            <span className="font-mono font-bold text-navy">{wo.id}</span>
            {wo.location && <> — {wo.location}</>}
          </p>
        </div>

        <div className="bg-slate-50 border border-slate-200 rounded-lg p-3 text-sm text-slate-700 space-y-1.5">
          <p className="font-semibold text-slate-800">This will permanently remove:</p>
          <ul className="list-disc list-inside text-xs space-y-0.5 pl-1">
            <li>The Work Order Tracker row</li>
            <li>All Marking Items rows keyed to this WO</li>
          </ul>
          <p className="font-semibold text-slate-800 pt-1">Preserved (audit trail):</p>
          <ul className="list-disc list-inside text-xs space-y-0.5 pl-1">
            <li>Work Day Log entries</li>
            <li>Sign-In / Payroll history</li>
            <li>Doc Lifecycle Log rows</li>
            <li>The WO's Drive archive folder &amp; any filed documents</li>
          </ul>
        </div>

        {wasWorked && (
          <div className="bg-amber-50 border border-amber-200 rounded-lg p-3 text-xs text-amber-800">
            <strong>Note:</strong> this WO has been worked on. Deleting won't
            retract any time-card, payroll, or completed-doc records.
          </div>
        )}

        <p className="text-xs text-slate-500 text-center font-semibold">
          This cannot be undone.
        </p>

        {error && (
          <p className="text-xs text-red-600 font-semibold text-center">{error}</p>
        )}

        <div className="flex flex-col gap-2 pt-1">
          <button
            type="button"
            disabled={submitting}
            onClick={doDelete}
            className="w-full py-3 rounded-xl font-bold text-sm bg-red-500 text-white
                       hover:bg-red-600 active:opacity-80 transition-all
                       disabled:opacity-40"
          >
            {submitting ? 'Deleting…' : `Delete WO ${wo.id}`}
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
