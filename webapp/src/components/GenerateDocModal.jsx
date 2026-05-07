import { useEffect, useState } from 'react'

/**
 * GenerateDocModal — blocking modal for document-generation actions.
 *
 * Replaces the transient bottom-corner toast that flashes for 6s and
 * disappears. This modal stays open until the user explicitly dismisses
 * a success or error state, and installs a `beforeunload` guard while
 * the request is in flight so an accidental tab-close prompts a warning.
 *
 * Lifecycle:
 *   idle       — show description + Cancel / Generate
 *   submitting — Cancel disabled, Generate replaced with spinner.
 *                Backdrop click + Esc disabled. beforeunload guard on.
 *   success    — green check + message + (optional) file list + OK
 *   error      — red X + reason + Try Again + OK
 *
 * Props:
 *   open          bool          — render or not
 *   title         string        — modal heading
 *   description   ReactNode     — preflight info shown in idle state
 *   confirmLabel  string?       — button text (default 'Generate')
 *   onConfirm     async fn      — runs the request. Resolve with
 *                                 `{ message, files? }` for success,
 *                                 throw `Error(message)` for failure.
 *   onClose       fn            — fires only when state allows dismiss
 *                                 (idle Cancel, success/error OK).
 */
export default function GenerateDocModal({
  open,
  title,
  description,
  confirmLabel = 'Generate',
  onConfirm,
  onClose,
}) {
  const [state,   setState]   = useState('idle')   // 'idle' | 'submitting' | 'success' | 'error'
  const [message, setMessage] = useState('')
  const [files,   setFiles]   = useState([])

  // Reset state every time the modal is (re)opened.
  useEffect(() => {
    if (open) {
      setState('idle')
      setMessage('')
      setFiles([])
    }
  }, [open])

  // beforeunload guard while submitting.
  useEffect(() => {
    if (state !== 'submitting') return
    const handler = (e) => { e.preventDefault(); e.returnValue = '' }
    window.addEventListener('beforeunload', handler)
    return () => window.removeEventListener('beforeunload', handler)
  }, [state])

  // Esc closes only when state allows dismissal.
  useEffect(() => {
    if (!open) return
    const handler = (e) => {
      if (e.key !== 'Escape') return
      if (state === 'submitting') return
      onClose?.()
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [open, state, onClose])

  if (!open) return null

  const handleConfirm = async () => {
    if (state === 'submitting') return
    setState('submitting')
    try {
      const result = await onConfirm()
      const msg = (result && result.message) || 'Done.'
      setMessage(msg)
      setFiles((result && result.files) || [])
      setState('success')
    } catch (err) {
      setMessage(err?.message || 'Something went wrong.')
      setState('error')
    }
  }

  const handleBackdrop = () => {
    if (state === 'submitting') return
    onClose?.()
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4"
      style={{ backgroundColor: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(2px)' }}
      onClick={handleBackdrop}
    >
      <div
        className="bg-white rounded-2xl shadow-2xl w-full max-w-md p-6 space-y-4"
        onClick={e => e.stopPropagation()}
      >
        {/* Icon */}
        <div className="text-4xl text-center">
          {state === 'success' ? (
            <span className="inline-flex items-center justify-center w-12 h-12 rounded-full bg-green-100 text-green-600">✓</span>
          ) : state === 'error' ? (
            <span className="inline-flex items-center justify-center w-12 h-12 rounded-full bg-red-100 text-red-600">✕</span>
          ) : state === 'submitting' ? (
            <span className="inline-block w-10 h-10 rounded-full border-4 border-slate-200 border-t-navy animate-spin" />
          ) : (
            <span>📄</span>
          )}
        </div>

        {/* Title + body */}
        <div className="text-center space-y-2">
          <h2 className="text-lg font-black text-navy">{title}</h2>
          {state === 'idle' && (
            <div className="text-slate-500 text-sm leading-relaxed">{description}</div>
          )}
          {state === 'submitting' && (
            <p className="text-slate-500 text-sm leading-relaxed">
              Working on it… don't close this tab.
            </p>
          )}
          {state === 'success' && (
            <div className="space-y-2">
              <p className="text-slate-700 text-sm leading-relaxed">{message}</p>
              {files.length > 0 && (
                <ul className="text-[11px] font-mono text-slate-500 bg-slate-50 rounded-lg p-2 text-left space-y-0.5 max-h-[120px] overflow-y-auto">
                  {files.map((f, i) => <li key={i} className="truncate">• {f}</li>)}
                </ul>
              )}
            </div>
          )}
          {state === 'error' && (
            <p className="text-slate-700 text-sm leading-relaxed whitespace-pre-line">{message}</p>
          )}
        </div>

        {/* Buttons */}
        <div className="flex flex-col gap-2 pt-1">
          {state === 'idle' && (
            <>
              <button
                onClick={handleConfirm}
                className="w-full py-3 rounded-xl font-bold text-sm bg-green-500 text-white hover:bg-green-600 active:opacity-80 transition-all"
              >
                {confirmLabel}
              </button>
              <button
                onClick={onClose}
                className="w-full py-3 rounded-xl font-bold text-sm bg-slate-100 text-slate-600 hover:bg-slate-200 transition-all"
              >
                Cancel
              </button>
            </>
          )}
          {state === 'submitting' && (
            <button
              disabled
              className="w-full py-3 rounded-xl font-bold text-sm bg-slate-100 text-slate-300 cursor-not-allowed"
            >
              Cancel (locked while running)
            </button>
          )}
          {state === 'success' && (
            <button
              onClick={onClose}
              className="w-full py-3 rounded-xl font-bold text-sm bg-navy text-white hover:opacity-90 active:opacity-80 transition-all"
            >
              OK
            </button>
          )}
          {state === 'error' && (
            <>
              <button
                onClick={() => setState('idle')}
                className="w-full py-3 rounded-xl font-bold text-sm bg-amber-500 text-white hover:bg-amber-600 active:opacity-80 transition-all"
              >
                Try Again
              </button>
              <button
                onClick={onClose}
                className="w-full py-3 rounded-xl font-bold text-sm bg-slate-100 text-slate-600 hover:bg-slate-200 transition-all"
              >
                Close
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  )
}
