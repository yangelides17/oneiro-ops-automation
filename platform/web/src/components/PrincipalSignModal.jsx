import { useEffect, useState } from 'react'
import SignaturePad from './SignaturePad'

// Mirrors _crop_to_ink in workers/fill_signin.py — crops a PNG data URL
// to its non-transparent bounding box (with an 8-px margin) so the
// server-side pdf-lib overlay doesn't scale a bunch of empty canvas.
// Returns a fresh PNG data URL; falls back to the input on any error.
async function cropPngToInk(dataUrl, marginPx = 8) {
  if (!dataUrl) return dataUrl
  try {
    const img = await new Promise((resolve, reject) => {
      const i = new Image()
      i.onload = () => resolve(i)
      i.onerror = reject
      i.src = dataUrl
    })
    const c   = document.createElement('canvas')
    c.width   = img.width
    c.height  = img.height
    const ctx = c.getContext('2d')
    ctx.drawImage(img, 0, 0)
    const { data, width, height } = ctx.getImageData(0, 0, c.width, c.height)
    let minX = width, minY = height, maxX = -1, maxY = -1
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        if (data[(y * width + x) * 4 + 3] > 0) {
          if (x < minX) minX = x
          if (y < minY) minY = y
          if (x > maxX) maxX = x
          if (y > maxY) maxY = y
        }
      }
    }
    if (maxX < 0) return dataUrl  // fully transparent → pass through
    minX = Math.max(0, minX - marginPx)
    minY = Math.max(0, minY - marginPx)
    maxX = Math.min(width  - 1, maxX + marginPx)
    maxY = Math.min(height - 1, maxY + marginPx)
    const cw = maxX - minX + 1
    const ch = maxY - minY + 1
    const out = document.createElement('canvas')
    out.width  = cw
    out.height = ch
    out.getContext('2d').drawImage(c, minX, minY, cw, ch, 0, 0, cw, ch)
    return out.toDataURL('image/png')
  } catch {
    return dataUrl
  }
}

/**
 * Opens before a Sign-In document gets approved. Collects the
 * principal's signature (base64 PNG), printed name, and title.
 * Submit button calls /api/approvals/:fileId/approve-signin which
 * patches the PDF via pdf-lib and moves it to ✅ Approved Docs.
 *
 * Props:
 *   filename   — shown in the header so the admin knows which doc they're signing off on
 *   onCancel   — close without action
 *   onSigned   — called after the approval succeeds; parent refreshes the list
 *                and auto-advances to the next item
 *   signUrl    — POST endpoint (defaults to /api/approvals/:fileId/approve-signin)
 */
export default function PrincipalSignModal({ filename, onCancel, onSigned, signUrl }) {
  const [signature, setSignature] = useState(null)
  const [name,      setName]      = useState('')
  const [title,     setTitle]     = useState('')
  const [submitting, setSubmitting] = useState(false)
  const [error, setError] = useState('')

  // Close on Escape
  useEffect(() => {
    const handler = (e) => { if (e.key === 'Escape' && !submitting) onCancel?.() }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [submitting, onCancel])

  const canSubmit = !!signature && !!name.trim() && !!title.trim() && !submitting

  const handleSubmit = async () => {
    setError('')
    if (!signature)    { setError('Signature is required.'); return }
    if (!name.trim())  { setError('Printed name is required.'); return }
    if (!title.trim()) { setError('Title / position is required.'); return }

    setSubmitting(true)
    try {
      // Crop to ink bbox before upload so the server-side overlay scales
      // the actual signature, not a bunch of empty canvas padding.
      const cropped = await cropPngToInk(signature)
      const res  = await fetch(signUrl, {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({
          signature_b64: cropped,
          name:          name.trim(),
          title:         title.trim(),
        })
      })
      const data = await res.json().catch(() => ({}))
      if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
      onSigned?.(data)
    } catch (err) {
      setError(err.message || 'Approval failed')
      setSubmitting(false)
    }
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4"
      style={{ backgroundColor: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(2px)' }}
      onClick={() => { if (!submitting) onCancel?.() }}
    >
      <div
        className="bg-white rounded-2xl shadow-2xl w-full max-w-lg p-5 space-y-4 max-h-[92vh] overflow-y-auto"
        onClick={e => e.stopPropagation()}
      >
        <div>
          <h2 className="text-lg font-black text-navy">Approve &amp; Sign</h2>
          <p className="text-[11px] text-slate-400 font-mono mt-0.5 truncate">{filename}</p>
          <p className="text-xs text-slate-500 mt-2 leading-snug">
            Sign below as the approving principal, then enter your printed
            name and title. Today&apos;s date is added automatically.
          </p>
        </div>

        {error && (
          <div className="bg-red-50 border border-red-200 text-red-700 text-xs px-3 py-2 rounded-lg">
            {error}
          </div>
        )}

        <div className="space-y-3">
          <SignaturePad
            label="Principal Signature"
            onChange={setSignature}
          />

          <div>
            <label className="field-label">Printed Name <span className="text-red-400 ml-0.5">*</span></label>
            <input
              type="text"
              value={name}
              autoCapitalize="words"
              placeholder="First Last"
              onChange={e => setName(e.target.value)}
              className="field-input"
            />
          </div>

          <div>
            <label className="field-label">Title / Position <span className="text-red-400 ml-0.5">*</span></label>
            <input
              type="text"
              value={title}
              autoCapitalize="words"
              placeholder="e.g. Principal / Owner / Operations Manager"
              onChange={e => setTitle(e.target.value)}
              className="field-input"
            />
          </div>
        </div>

        <div className="flex flex-col gap-2 pt-1">
          <button
            type="button"
            onClick={handleSubmit}
            disabled={!canSubmit}
            className="w-full py-3 rounded-xl font-bold text-sm
                       bg-green-600 text-white hover:bg-green-700 active:bg-green-800
                       disabled:opacity-50 disabled:cursor-not-allowed transition-all"
          >
            {submitting ? 'Signing & approving…' : 'Approve & Sign'}
          </button>
          <button
            type="button"
            onClick={onCancel}
            disabled={submitting}
            className="w-full py-3 rounded-xl font-bold text-sm
                       bg-slate-100 text-slate-600 hover:bg-slate-200
                       disabled:opacity-60 disabled:cursor-not-allowed transition-all"
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  )
}
