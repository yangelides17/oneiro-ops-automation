import { useEffect, useRef, useState } from 'react'
import SignaturePad from './SignaturePad'
import { cropPngToInk, signatureFromPhoto } from '../lib/signatureImage'

const DRAW = '__draw__'
const LAST_SIGNER_KEY = 'oneiro.cpSignAs'

// Today as YYYY-MM-DD in the business timezone — the browser may not
// be in New York, and toISOString() is UTC (rolls over at 20:00 ET).
function todayET() {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: 'America/New_York', year: 'numeric', month: '2-digit', day: '2-digit',
  }).format(new Date())
}

function readLastSigner() {
  try { return localStorage.getItem(LAST_SIGNER_KEY) || '' } catch { return '' }
}
function writeLastSigner(id) {
  try { localStorage.setItem(LAST_SIGNER_KEY, id) } catch { /* private mode */ }
}

const pngSrc = (b64) => (b64 ? `data:image/png;base64,${b64}` : '')

/**
 * Approve & Sign for Certified Payroll. Signs as a saved principal
 * ("Sign As …" — signature image stored via /api/principal-signers) or
 * with a signature drawn here, stamps printed name + title + the chosen
 * sign date, then POSTs to /api/approvals/:fileId/approve-cert-payroll.
 *
 * Props:
 *   fileId    — the pending CP's Drive file id
 *   filename  — shown in the header
 *   onCancel  — close without action
 *   onSigned  — called after the approval succeeds
 */
export default function CertPayrollSignModal({ fileId, filename, onCancel, onSigned }) {
  const [signers, setSigners]     = useState(null)   // null = loading
  const [loadError, setLoadError] = useState('')
  const [signAs, setSignAs]       = useState('')     // signer id | DRAW
  const [drawn, setDrawn]         = useState(null)
  const [name, setName]           = useState('')
  const [title, setTitle]         = useState('')
  const [signDate, setSignDate]   = useState(todayET())
  const [editor, setEditor]       = useState(null)   // null | { id?, name, title }
  const [confirmRemove, setConfirmRemove] = useState(false)
  const [submitting, setSubmitting] = useState(false)
  const [error, setError]         = useState('')

  const busy = submitting || !!editor

  const loadSigners = async (preferId) => {
    setLoadError('')
    try {
      const res  = await fetch('/api/principal-signers')
      const data = await res.json().catch(() => ({}))
      if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
      const list = data.signers || []
      setSigners(list)
      const pick = [preferId, readLastSigner()].find(id => id && list.some(s => s.id === id))
                || list[0]?.id || DRAW
      choose(pick, list)
      if (!list.length) setEditor({ name: '', title: '' })
    } catch (err) {
      setSigners([])
      setLoadError(err.message || 'Couldn’t load saved signers')
      choose(DRAW, [])
    }
  }

  useEffect(() => { loadSigners() }, [])  // eslint-disable-line react-hooks/exhaustive-deps

  // Close on Escape
  useEffect(() => {
    const handler = (e) => { if (e.key === 'Escape' && !submitting) onCancel?.() }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [submitting, onCancel])

  // Selecting a signer pre-fills name + title (still editable — e.g. a
  // one-off title change shouldn't require editing the saved signer).
  const choose = (id, list = signers || []) => {
    setSignAs(id)
    setConfirmRemove(false)
    const s = list.find(x => x.id === id)
    if (s) {
      setName(s.name)
      setTitle(s.title)
      writeLastSigner(id)
    }
  }

  const current = (signers || []).find(s => s.id === signAs)
  const hasSignature = signAs === DRAW ? !!drawn : !!current?.png_b64
  const canSubmit = hasSignature && !!name.trim() && !!title.trim() && !!signDate && !busy

  const removeSigner = async () => {
    if (!current) return
    setError('')
    try {
      const res  = await fetch(`/api/principal-signers/${encodeURIComponent(current.id)}`,
                               { method: 'DELETE' })
      const data = await res.json().catch(() => ({}))
      if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
      setName(''); setTitle('')
      await loadSigners()
    } catch (err) {
      setError(err.message || 'Remove failed')
    }
  }

  const handleSubmit = async () => {
    setError('')
    if (!hasSignature) { setError('Signature is required.'); return }
    if (!name.trim())  { setError('Printed name is required.'); return }
    if (!title.trim()) { setError('Title / position is required.'); return }
    if (!signDate)     { setError('Sign date is required.'); return }

    setSubmitting(true)
    try {
      const payload = { name: name.trim(), title: title.trim(), sign_date: signDate }
      if (signAs === DRAW) payload.signature_b64 = await cropPngToInk(drawn)
      else                 payload.signer_id     = signAs
      const res  = await fetch(
        `/api/approvals/${encodeURIComponent(fileId)}/approve-cert-payroll`, {
          method:  'POST',
          headers: { 'Content-Type': 'application/json' },
          body:    JSON.stringify(payload),
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
        </div>

        {error && (
          <div className="bg-red-50 border border-red-200 text-red-700 text-xs px-3 py-2 rounded-lg">
            {error}
          </div>
        )}

        {editor ? (
          <SignerEditor
            initial={editor}
            existing={editor.id ? (signers || []).find(s => s.id === editor.id) : null}
            firstSigner={!signers?.length}
            onCancel={() => setEditor(null)}
            onSaved={async (saved) => { setEditor(null); await loadSigners(saved.id) }}
          />
        ) : (
          <div className="space-y-3">
            <div>
              <label className="field-label">Sign As</label>
              {signers === null ? (
                <p className="text-xs text-slate-400 py-2">Loading saved signers…</p>
              ) : (
                <div className="flex gap-2">
                  <select
                    value={signAs}
                    onChange={e => choose(e.target.value)}
                    className="field-input flex-1"
                  >
                    {signers.map(s => (
                      <option key={s.id} value={s.id}>{s.name} — {s.title}</option>
                    ))}
                    <option value={DRAW}>Draw a signature instead</option>
                  </select>
                  {!loadError && (
                    <button
                      type="button"
                      onClick={() => setEditor({ name: '', title: '' })}
                      className="text-xs font-bold px-3 rounded-lg bg-slate-100 text-slate-600
                                 hover:bg-slate-200 whitespace-nowrap"
                    >
                      + Add signer
                    </button>
                  )}
                </div>
              )}
              {loadError && (
                <p className="text-[11px] text-amber-700 mt-1">
                  Saved signers unavailable ({loadError}). You can still draw a signature.
                </p>
              )}
            </div>

            {signAs === DRAW ? (
              <SignaturePad label="Principal Signature" onChange={setDrawn} />
            ) : current && (
              <div>
                <div className="border border-slate-200 rounded-lg bg-white h-24 flex items-center justify-center p-2">
                  {current.png_b64
                    ? <img src={pngSrc(current.png_b64)} alt={`${current.name} signature`}
                           className="max-h-full max-w-full object-contain" />
                    : <span className="text-xs text-red-600">{current.error || 'Signature image missing'}</span>}
                </div>
                <div className="flex justify-end gap-3 mt-1 text-[11px] font-semibold">
                  <button type="button" className="text-slate-500 hover:text-navy"
                          onClick={() => setEditor({ id: current.id, name: current.name, title: current.title })}>
                    Edit signer
                  </button>
                  {confirmRemove ? (
                    <span className="text-red-600">
                      Remove {current.name}?{' '}
                      <button type="button" className="underline" onClick={removeSigner}>Yes</button>{' '}
                      <button type="button" className="underline text-slate-500"
                              onClick={() => setConfirmRemove(false)}>No</button>
                    </span>
                  ) : (
                    <button type="button" className="text-slate-500 hover:text-red-600"
                            onClick={() => setConfirmRemove(true)}>
                      Remove
                    </button>
                  )}
                </div>
              </div>
            )}

            <div>
              <label className="field-label">Printed Name <span className="text-red-400 ml-0.5">*</span></label>
              <input type="text" value={name} autoCapitalize="words" placeholder="First Last"
                     onChange={e => setName(e.target.value)} className="field-input" />
            </div>

            <div>
              <label className="field-label">Title / Position <span className="text-red-400 ml-0.5">*</span></label>
              <input type="text" value={title} autoCapitalize="words"
                     placeholder="e.g. Principal / Owner"
                     onChange={e => setTitle(e.target.value)} className="field-input" />
            </div>

            <div>
              <label className="field-label">Sign Date <span className="text-red-400 ml-0.5">*</span></label>
              <input type="date" value={signDate} required
                     onChange={e => setSignDate(e.target.value)} className="field-input" />
              {signDate !== todayET() && (
                <button type="button" onClick={() => setSignDate(todayET())}
                        className="text-[11px] font-semibold text-slate-500 hover:text-navy mt-1">
                  Reset to today
                </button>
              )}
            </div>
          </div>
        )}

        {!editor && (
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
        )}
      </div>
    </div>
  )
}

/**
 * Add / edit a saved signer: name, title, and a signature image either
 * uploaded (photo or scan of a pen signature → background removed) or
 * drawn. When editing, leaving the image untouched keeps the old one.
 */
function SignerEditor({ initial, existing, firstSigner, onCancel, onSaved }) {
  const [name, setName]   = useState(initial.name || '')
  const [title, setTitle] = useState(initial.title || '')
  const [mode, setMode]   = useState('upload')    // 'upload' | 'draw'
  const [image, setImage] = useState(null)        // new PNG data URL
  const [processing, setProcessing] = useState(false)
  const [saving, setSaving] = useState(false)
  const [error, setError]   = useState('')
  const fileRef = useRef(null)

  const isEdit = !!initial.id
  const preview = image || (isEdit ? pngSrc(existing?.png_b64) : '')
  const canSave = !!name.trim() && !!title.trim() && (isEdit || !!image) && !saving && !processing

  const onFile = async (e) => {
    const file = e.target.files?.[0]
    e.target.value = ''
    if (!file) return
    setError('')
    setProcessing(true)
    try {
      setImage(await signatureFromPhoto(file))
    } catch (err) {
      setError(err.message || 'Couldn’t process that image')
    } finally {
      setProcessing(false)
    }
  }

  const save = async () => {
    setError('')
    setSaving(true)
    try {
      const body = { name: name.trim(), title: title.trim() }
      if (isEdit) body.id = initial.id
      if (image)  body.png_b64 = mode === 'draw' ? await cropPngToInk(image) : image
      const res  = await fetch('/api/principal-signers', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify(body),
      })
      const data = await res.json().catch(() => ({}))
      if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
      onSaved?.(data.signer)
    } catch (err) {
      setError(err.message || 'Save failed')
      setSaving(false)
    }
  }

  const tab = (value, label) => (
    <button
      type="button"
      onClick={() => { setMode(value); setImage(null) }}
      className={`flex-1 text-xs font-bold py-1.5 rounded-md transition-colors ${
        mode === value ? 'bg-white text-navy shadow-sm' : 'text-slate-500 hover:text-navy'}`}
    >
      {label}
    </button>
  )

  return (
    <div className="space-y-3 border border-slate-200 rounded-xl p-3 bg-slate-50">
      <p className="text-sm font-bold text-navy">{isEdit ? 'Edit signer' : 'Add a signer'}</p>
      {!isEdit && firstSigner && (
        <p className="text-xs text-slate-500 leading-snug">
          No saved signers yet. Save a principal’s signature once and sign
          future payrolls in one click.
        </p>
      )}

      {error && (
        <div className="bg-red-50 border border-red-200 text-red-700 text-xs px-3 py-2 rounded-lg">
          {error}
        </div>
      )}

      <div>
        <label className="field-label">Printed Name <span className="text-red-400 ml-0.5">*</span></label>
        <input type="text" value={name} autoCapitalize="words" placeholder="First Last"
               onChange={e => setName(e.target.value)} className="field-input" />
      </div>
      <div>
        <label className="field-label">Title / Position <span className="text-red-400 ml-0.5">*</span></label>
        <input type="text" value={title} autoCapitalize="words" placeholder="e.g. Principal / Owner"
               onChange={e => setTitle(e.target.value)} className="field-input" />
      </div>

      <div>
        <label className="field-label">
          Signature {!isEdit && <span className="text-red-400 ml-0.5">*</span>}
        </label>
        <div className="flex gap-1 bg-slate-200 rounded-lg p-1 mb-2">
          {tab('upload', 'Upload image')}
          {tab('draw', 'Draw')}
        </div>

        {mode === 'upload' ? (
          <>
            <input ref={fileRef} type="file" accept="image/*,.heic,.heif"
                   onChange={onFile} className="hidden" />
            <button
              type="button"
              onClick={() => fileRef.current?.click()}
              disabled={processing}
              className="w-full border-2 border-dashed border-slate-300 rounded-lg bg-white
                         h-24 flex items-center justify-center p-2 hover:border-slate-400"
            >
              {processing ? (
                <span className="text-xs text-slate-400">Processing…</span>
              ) : preview ? (
                <img src={preview} alt="Signature preview" className="max-h-full max-w-full object-contain" />
              ) : (
                <span className="text-xs text-slate-500">
                  Choose a photo or scan of the signature (dark pen on white paper)
                </span>
              )}
            </button>
            <p className="text-[11px] text-slate-400 mt-1">
              {preview
                ? (image ? 'Background removed — tap to choose a different image.'
                         : 'Current signature — tap to replace it.')
                : 'The paper background is removed automatically.'}
            </p>
          </>
        ) : (
          <SignaturePad label="" onChange={setImage} />
        )}
      </div>

      <div className="flex gap-2">
        <button
          type="button"
          onClick={save}
          disabled={!canSave}
          className="flex-1 py-2.5 rounded-lg font-bold text-sm bg-navy text-white
                     hover:opacity-90 disabled:opacity-50 disabled:cursor-not-allowed"
        >
          {saving ? 'Saving…' : 'Save signer'}
        </button>
        <button
          type="button"
          onClick={onCancel}
          disabled={saving}
          className="flex-1 py-2.5 rounded-lg font-bold text-sm bg-slate-100 text-slate-600
                     hover:bg-slate-200 disabled:opacity-60"
        >
          {firstSigner && !isEdit ? 'Skip — draw instead' : 'Cancel'}
        </button>
      </div>
    </div>
  )
}
