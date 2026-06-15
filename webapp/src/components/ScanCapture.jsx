import { useEffect, useRef, useState } from 'react'
import {
  loadScanner, imageToCanvas, detectCorners, warpAndClean,
} from '../lib/docScanner'
import { imagesToPdf, normalizeHeic } from '../lib/imagesToPdf'

// In-app document scanner for sign-in sheets. The crew taps "Take
// Photos" → the phone's native camera opens → each page is auto-cropped
// & de-skewed (with draggable corner handles to fine-tune) → all pages
// are cleaned to a crisp "scanned" look and assembled into one PDF,
// which is handed to onScanned() to flow through the existing upload/OCR
// path. onCancel() backs out.
//
// Corners are stored in SOURCE-image coordinates; the overlay scales
// them to the displayed size.

const CORNER_KEYS = ['topLeftCorner', 'topRightCorner', 'bottomRightCorner', 'bottomLeftCorner']
const DISPLAY_W = 320
const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36)

export default function ScanCapture({ onScanned, onCancel }) {
  const [pages, setPages]   = useState([])   // { id, canvas, width, height, dataUrl, corners }
  const [activeIdx, setActiveIdx] = useState(0)
  const [busy, setBusy]     = useState('')   // '' | 'reading' | 'building'
  const [error, setError]   = useState('')
  const fileInputRef = useRef(null)

  // Warm the OpenCV runtime as soon as the scanner opens.
  useEffect(() => { loadScanner().catch(() => {}) }, [])

  const addFiles = async (fileList) => {
    const files = Array.from(fileList || [])
    if (!files.length) return
    setError(''); setBusy('reading')
    try {
      const next = []
      for (const raw of files) {
        const file = await normalizeHeic(raw)
        const { canvas, width, height } = await imageToCanvas(file)
        const corners = await detectCorners(canvas)
        next.push({ id: uid(), canvas, width, height, dataUrl: canvas.toDataURL('image/jpeg', 0.9), corners })
      }
      setPages(prev => {
        const merged = [...prev, ...next]
        setActiveIdx(merged.length - next.length)   // jump to first new page
        return merged
      })
    } catch (err) {
      setError(err.message || 'Could not read photo')
    } finally {
      setBusy('')
    }
  }

  const removePage = (id) => {
    setPages(prev => {
      const merged = prev.filter(p => p.id !== id)
      setActiveIdx(i => Math.max(0, Math.min(i, merged.length - 1)))
      return merged
    })
  }

  const setCorner = (key, pt) => {
    setPages(prev => prev.map((p, i) =>
      i === activeIdx ? { ...p, corners: { ...p.corners, [key]: pt } } : p))
  }

  const buildPdf = async () => {
    if (!pages.length || busy) return
    setError(''); setBusy('building')
    try {
      const blobs = []
      for (const p of pages) blobs.push(await warpAndClean(p.canvas, p.corners))
      const bytes = await imagesToPdf(blobs)
      const file = new File([bytes], `scan-${Date.now()}.pdf`, { type: 'application/pdf' })
      onScanned(file)
    } catch (err) {
      setError(err.message || 'Could not build the scan')
      setBusy('')
    }
  }

  const active = pages[activeIdx] || null

  return (
    <div className="card p-4 space-y-3">
      <div className="flex items-center justify-between">
        <p className="section-label !mb-0">Scan Sheet</p>
        <button type="button" onClick={onCancel}
          className="text-[11px] text-slate-500 font-semibold hover:underline">
          Cancel
        </button>
      </div>

      <input
        ref={fileInputRef}
        type="file"
        accept="image/*"
        capture="environment"
        multiple
        className="hidden"
        onChange={e => { addFiles(e.target.files); e.target.value = '' }}
      />

      {pages.length === 0 ? (
        <button
          type="button"
          onClick={() => fileInputRef.current?.click()}
          disabled={busy === 'reading'}
          className="w-full border-2 border-dashed border-slate-300 rounded-xl p-8 text-center
                     hover:border-navy transition-colors disabled:opacity-60">
          <div className="text-3xl mb-1">📷</div>
          <p className="text-sm font-semibold text-navy">
            {busy === 'reading' ? 'Reading…' : 'Take photos of the sheet'}
          </p>
          <p className="text-[12px] text-slate-500 mt-1">
            Capture each page — edges are auto-detected and cleaned up.
          </p>
        </button>
      ) : (
        <>
          {/* Corner-adjust view for the active page */}
          {active && (
            <CornerAdjuster
              page={active}
              onCorner={setCorner}
            />
          )}

          {/* Page strip */}
          <div className="flex items-center gap-2 overflow-x-auto py-1">
            {pages.map((p, i) => (
              <div key={p.id} className="relative flex-shrink-0">
                <button
                  type="button"
                  onClick={() => setActiveIdx(i)}
                  className={`block rounded-lg overflow-hidden border-2 ${i === activeIdx ? 'border-navy' : 'border-slate-200'}`}>
                  <img src={p.dataUrl} alt={`page ${i + 1}`} className="h-16 w-12 object-cover" />
                </button>
                <button
                  type="button"
                  onClick={() => removePage(p.id)}
                  className="absolute -top-1.5 -right-1.5 bg-red-500 text-white rounded-full w-5 h-5
                             text-xs leading-none flex items-center justify-center shadow">
                  ×
                </button>
              </div>
            ))}
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              disabled={busy === 'reading'}
              className="flex-shrink-0 h-16 w-12 rounded-lg border-2 border-dashed border-slate-300
                         text-navy text-xl hover:border-navy disabled:opacity-60">
              +
            </button>
          </div>

          <p className="text-[11px] text-slate-400">
            Drag the corner handles to match the sheet edges if the auto-detection is off.
          </p>

          <button
            type="button"
            onClick={buildPdf}
            disabled={!!busy}
            className="btn-primary w-full">
            {busy === 'building' ? 'Building scan…' : `Use ${pages.length} page${pages.length > 1 ? 's' : ''}`}
          </button>
        </>
      )}

      {error && <p className="text-[12px] text-red-600">{error}</p>}
    </div>
  )
}

// ── Corner adjuster ───────────────────────────────────────────
// Renders the page at a fixed display width with a draggable quad. All
// math is done in source coordinates; the SVG just scales for display.
function CornerAdjuster({ page, onCorner }) {
  const scale = DISPLAY_W / page.width
  const dispH = Math.round(page.height * scale)
  const svgRef = useRef(null)
  const dragKey = useRef(null)

  const toSource = (clientX, clientY) => {
    const rect = svgRef.current.getBoundingClientRect()
    const x = (clientX - rect.left) / scale
    const y = (clientY - rect.top) / scale
    return {
      x: Math.max(0, Math.min(page.width, x)),
      y: Math.max(0, Math.min(page.height, y)),
    }
  }

  useEffect(() => {
    const move = (e) => {
      if (!dragKey.current) return
      const t = e.touches ? e.touches[0] : e
      onCorner(dragKey.current, toSource(t.clientX, t.clientY))
    }
    const up = () => { dragKey.current = null }
    window.addEventListener('pointermove', move)
    window.addEventListener('pointerup', up)
    return () => {
      window.removeEventListener('pointermove', move)
      window.removeEventListener('pointerup', up)
    }
  }, [page, scale])

  const c = page.corners
  const poly = CORNER_KEYS.map(k => `${c[k].x * scale},${c[k].y * scale}`).join(' ')

  return (
    <div className="relative mx-auto bg-slate-100 rounded-lg overflow-hidden"
         style={{ width: DISPLAY_W, height: dispH }}>
      <img src={page.dataUrl} alt="page" width={DISPLAY_W} height={dispH} draggable={false} />
      <svg
        ref={svgRef}
        className="absolute inset-0 touch-none"
        width={DISPLAY_W}
        height={dispH}
      >
        <polygon points={poly} fill="rgba(37,99,235,0.12)" stroke="#2563eb" strokeWidth="2" />
        {CORNER_KEYS.map(k => (
          <circle
            key={k}
            cx={c[k].x * scale}
            cy={c[k].y * scale}
            r="9"
            fill="#ffffff"
            stroke="#2563eb"
            strokeWidth="2"
            style={{ cursor: 'grab' }}
            onPointerDown={(e) => { e.preventDefault(); dragKey.current = k }}
          />
        ))}
      </svg>
    </div>
  )
}
