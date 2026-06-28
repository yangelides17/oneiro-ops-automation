import { useEffect, useRef, useState } from 'react'
import ScanCapture from './ScanCapture'

// Replace a pending-approval PDF with a new signed/rescanned version.
// Mirrors the Sign-In page's "Upload / Scan" affordance but WITHOUT the
// Claude Vision parse — we just forward the raw PDF File. Sign-in docs
// get both "Scan with camera" and "Choose PDF"; other doc types get
// "Choose PDF" only (allowScan=false).
//
// Props:
//   fileId, filename — the pending doc being replaced
//   allowScan        — show the camera-scan option (sign-in only)
//   onClose()        — dismiss
//   onReuploaded(newFileId) — fired after a successful replace

const MAX_BYTES = 20 * 1024 * 1024   // matches the server multer limit

export default function ReuploadModal({ fileId, filename, allowScan, onClose, onReuploaded }) {
  const [mode, setMode]     = useState('choose')   // 'choose' | 'scanning'
  const [status, setStatus] = useState('idle')     // 'idle' | 'uploading' | 'error'
  const [error, setError]   = useState('')
  const pdfInputRef = useRef(null)

  // Close on Escape (but not mid-upload).
  useEffect(() => {
    const handler = (e) => { if (e.key === 'Escape' && status !== 'uploading') onClose?.() }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [onClose, status])

  const uploadFile = async (file) => {
    if (!file) return
    if (file.size > MAX_BYTES) {
      setMode('choose'); setStatus('error')
      setError(`File is too large (${(file.size / 1e6).toFixed(1)} MB). Max is 20 MB.`)
      return
    }
    setMode('choose'); setStatus('uploading'); setError('')
    try {
      const fd = new FormData()
      fd.append('file', file)
      const res = await fetch(`/api/approvals/${encodeURIComponent(fileId)}/reupload`, {
        method: 'POST', body: fd,
      })
      const json = await res.json()
      if (!res.ok || json.error) throw new Error(json.error || `HTTP ${res.status}`)
      onReuploaded(json.file_id)
    } catch (err) {
      setStatus('error')
      setError(err.message || 'Reupload failed')
    }
  }

  const onPickPdf = (e) => {
    const f = e.target.files?.[0]
    e.target.value = ''   // allow re-picking the same file
    if (!f) return
    if (f.type && f.type !== 'application/pdf') {
      setStatus('error'); setError('Please choose a PDF file.')
      return
    }
    uploadFile(f)
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4"
      style={{ backgroundColor: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(2px)' }}
      onClick={() => { if (status !== 'uploading') onClose?.() }}
    >
      <div
        className="bg-white rounded-2xl shadow-2xl w-full max-w-md p-5 space-y-4 max-h-[90vh] overflow-y-auto"
        onClick={e => e.stopPropagation()}
      >
        <div className="flex items-start justify-between gap-3">
          <div className="min-w-0">
            <h2 className="text-lg font-black text-navy">Reupload Document</h2>
            <p className="text-[12px] text-slate-500 truncate">Replaces: {filename}</p>
          </div>
          <button
            type="button"
            onClick={() => { if (status !== 'uploading') onClose?.() }}
            className="text-slate-400 hover:text-slate-600 text-xl leading-none flex-shrink-0">
            ×
          </button>
        </div>

        {status === 'uploading' ? (
          <div className="flex items-center gap-3 py-8 justify-center">
            <div className="w-5 h-5 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
            <span className="text-sm text-slate-600">Uploading replacement…</span>
          </div>
        ) : mode === 'scanning' ? (
          <ScanCapture onScanned={uploadFile} onCancel={() => setMode('choose')} />
        ) : (
          <>
            <div className={`grid grid-cols-1 ${allowScan ? 'sm:grid-cols-2' : ''} gap-3`}>
              {allowScan && (
                <button
                  type="button"
                  onClick={() => { setStatus('idle'); setError(''); setMode('scanning') }}
                  className="border-2 border-dashed border-slate-300 rounded-xl p-6 text-center
                             hover:border-navy transition-colors">
                  <div className="text-3xl mb-1">📷</div>
                  <p className="text-sm font-semibold text-navy">Scan with camera</p>
                  <p className="text-[12px] text-slate-500 mt-1">
                    Photograph the signed sheet — cropped &amp; cleaned
                  </p>
                </button>
              )}
              <label className="block border-2 border-dashed border-slate-300 rounded-xl p-6 text-center cursor-pointer hover:border-navy transition-colors">
                <input
                  ref={pdfInputRef}
                  type="file"
                  accept="application/pdf"
                  className="hidden"
                  onChange={onPickPdf}
                />
                <div className="text-3xl mb-1">📄</div>
                <p className="text-sm font-semibold text-navy">Choose PDF</p>
                <p className="text-[12px] text-slate-500 mt-1">
                  Signed PDF from your device (≤ 20 MB)
                </p>
              </label>
            </div>
            {status === 'error' && error && (
              <p className="text-[12px] text-red-600">{error}</p>
            )}
            <p className="text-[11px] text-slate-400">
              The replacement keeps the same filename and stays in the review queue for approval.
            </p>
          </>
        )}
      </div>
    </div>
  )
}
