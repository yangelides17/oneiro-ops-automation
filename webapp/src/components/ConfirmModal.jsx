import { useEffect } from 'react'

/**
 * ConfirmModal — simple blocking confirmation dialog.
 *
 * Props:
 *   title        — modal heading
 *   message      — body text
 *   confirmLabel — text on the confirm button (default: "Confirm")
 *   cancelLabel  — text on the cancel button (default: "Cancel")
 *   danger       — if true, confirm button is red (default: false)
 *   onConfirm    — called when user clicks confirm
 *   onCancel     — called when user clicks cancel or the backdrop
 */
export default function ConfirmModal({
  title,
  message,
  confirmLabel = 'Confirm',
  cancelLabel  = 'Cancel',
  danger       = false,
  onConfirm,
  onCancel
}) {
  // Close on Escape
  useEffect(() => {
    const handler = (e) => { if (e.key === 'Escape') onCancel?.() }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [onCancel])

  return (
    /* Backdrop */
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4"
      style={{ backgroundColor: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(2px)' }}
      onClick={onCancel}
    >
      {/* Panel — stop clicks propagating to backdrop */}
      <div
        className="bg-white rounded-2xl shadow-2xl w-full max-w-sm p-6 space-y-4"
        onClick={e => e.stopPropagation()}
      >
        {/* Icon */}
        <div className="text-3xl text-center">⚠️</div>

        {/* Text */}
        <div className="text-center space-y-1.5">
          <h2 className="text-lg font-black text-navy">{title}</h2>
          <p className="text-slate-500 text-sm leading-relaxed">{message}</p>
        </div>

        {/* Buttons */}
        <div className="flex flex-col gap-2 pt-1">
          <button
            onClick={onConfirm}
            className={`w-full py-3 rounded-xl font-bold text-sm transition-all
              ${danger
                ? 'bg-red-500 text-white hover:bg-red-600 active:opacity-80'
                : 'bg-navy text-white hover:opacity-90 active:opacity-80'
              }`}
          >
            {confirmLabel}
          </button>
          <button
            onClick={onCancel}
            className="w-full py-3 rounded-xl font-bold text-sm bg-slate-100
                       text-slate-600 hover:bg-slate-200 transition-all"
          >
            {cancelLabel}
          </button>
        </div>
      </div>
    </div>
  )
}
