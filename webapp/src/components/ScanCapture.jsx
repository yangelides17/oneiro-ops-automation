import { useEffect, useRef, useState } from 'react'
import {
  loadScanner, onScannerProgress, imageToCanvas, detectCorners,
  warpAndClean, canvasToJpeg,
} from '../lib/docScanner'
import { imagesToPdf, normalizeHeic } from '../lib/imagesToPdf'

// In-app document scanner for sign-in sheets. The crew taps "Take
// Photos" → the phone's native camera opens → each page is auto-cropped
// & de-skewed (with draggable corner handles to fine-tune) → all pages
// are cleaned to a crisp "scanned" look and assembled into one PDF,
// which is handed to onScanned() to flow through the existing upload/OCR
// path. onCancel() backs out.
//
// The first scan downloads OpenCV.js (~9 MB) — we show a progress bar
// for that and log each step to the console ([docScanner]) so a stuck
// load is diagnosable. If OpenCV can't load at all, the user can still
// scan "without auto-crop" (pages uploaded as photographed) or fall
// back to choosing a PDF.

const CORNER_KEYS = ['topLeftCorner', 'topRightCorner', 'bottomRightCorner', 'bottomLeftCorner']
const DISPLAY_W = 320
const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36)

export default function ScanCapture({ onScanned, onCancel }) {
  const [pages, setPages]   = useState([])   // { id, canvas, width, height, dataUrl, corners|null }
  const [activeIdx, setActiveIdx] = useState(0)
  const [busy, setBusy]     = useState('')   // '' | 'reading' | 'building'
  const [error, setError]   = useState('')
  const [prog, setProg]     = useState({ phase: 'idle', ratio: 0, loaded: 0, indeterminate: false, error: '' })
  const [rawMode, setRawMode] = useState(false)   // skip OpenCV (no auto-crop)
  const fileInputRef = useRef(null)

  const ready  = prog.phase === 'ready'
  const failed = prog.phase === 'error'
  const useScanner = ready && !rawMode

  // Subscribe to load progress + kick the scanner load on mount.
  useEffect(() => {
    const off = onScannerProgress(setProg)
    loadScanner().catch(() => {})   // errors surface via progress phase
    return off
  }, [])

  const retryLoad = () => { setProg({ phase: 'downloading', ratio: 0, indeterminate: true, error: '' }); loadScanner().catch(() => {}) }

  const addFiles = async (fileList) => {
    const files = Array.from(fileList || [])
    if (!files.length) return
    setError(''); setBusy('reading')
    try {
      const next = []
      for (const raw of files) {
        const file = await normalizeHeic(raw)
        const { canvas, width, height } = await imageToCanvas(file)
        const corners = useScanner ? await detectCorners(canvas) : null
        next.push({ id: uid(), canvas, width, height, dataUrl: canvas.toDataURL('image/jpeg', 0.9), corners })
      }
      setPages(prev => {
        const merged = [...prev, ...next]
        setActiveIdx(merged.length - next.length)
        return merged
      })
    } catch (err) {
      console.warn('[ScanCapture] addFiles failed', err)
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
      i === activeIdx && p.corners ? { ...p, corners: { ...p.corners, [key]: pt } } : p))
  }

  const buildPdf = async () => {
    if (!pages.length || busy) return
    setError(''); setBusy('building')
    try {
      const blobs = []
      for (const p of pages) {
        blobs.push(p.corners ? await warpAndClean(p.canvas, p.corners) : await canvasToJpeg(p.canvas))
      }
      const bytes = await imagesToPdf(blobs)
      const file = new File([bytes], `scan-${Date.now()}.pdf`, { type: 'application/pdf' })
      onScanned(file)
    } catch (err) {
      console.warn('[ScanCapture] buildPdf failed', err)
      setError(err.message || 'Could not build the scan')
      setBusy('')
    }
  }

  const active = pages[activeIdx] || null
  // Capture is allowed once the scanner is ready, or in raw mode.
  const canCapture = useScanner || rawMode

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

      {/* Loader / error gating before capture is available */}
      {!canCapture && !failed && <LoadProgress prog={prog} onSkip={() => setRawMode(true)} />}

      {!canCapture && failed && (
        <div className="space-y-2">
          <p className="text-sm font-semibold text-red-600">Couldn’t load the scanner.</p>
          <p className="text-[12px] text-slate-500 break-words">
            {prog.error || 'The scanner components failed to download.'} Check your connection,
            then retry — or continue without auto-crop (pages upload as photographed), or use
            “Choose PDF” instead.
          </p>
          <div className="flex flex-wrap gap-2">
            <button type="button" onClick={retryLoad} className="btn-outline !py-1.5 text-sm">Retry</button>
            <button type="button" onClick={() => setRawMode(true)} className="btn-outline !py-1.5 text-sm">
              Continue without auto-crop
            </button>
          </div>
        </div>
      )}

      {/* Capture + staging */}
      {canCapture && (
        rawMode && (
          <div className="rounded-lg bg-amber-50 border border-amber-200 px-3 py-1.5 text-[11px] text-amber-800">
            Auto-crop is off — pages will be uploaded as photographed.
          </div>
        )
      )}

      {canCapture && pages.length === 0 && (
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
            {useScanner
              ? 'Capture each page — edges are auto-detected and cleaned up.'
              : 'Capture each page.'}
          </p>
        </button>
      )}

      {canCapture && pages.length > 0 && (
        <>
          {/* Corner-adjust view (auto mode) or plain preview (raw mode) */}
          {active && (active.corners
            ? <CornerAdjuster page={active} onCorner={setCorner} />
            : <div className="mx-auto bg-slate-100 rounded-lg overflow-hidden" style={{ width: DISPLAY_W }}>
                <img src={active.dataUrl} alt="page" width={DISPLAY_W} />
              </div>)}

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

          {useScanner && (
            <p className="text-[11px] text-slate-400">
              Drag the corner handles to match the sheet edges if the auto-detection is off.
            </p>
          )}

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

// ── Load progress indicator ───────────────────────────────────
function LoadProgress({ prog, onSkip }) {
  const pct = Math.round((prog.ratio || 0) * 100)
  const downloading  = prog.phase === 'downloading'
  const initializing = prog.phase === 'initializing'
  const showBar = downloading && !prog.indeterminate
  return (
    <div className="space-y-2 py-2">
      <div className="flex items-center gap-2 text-sm text-slate-600">
        <span className="w-4 h-4 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
        <span>
          {initializing ? 'Starting scanner…' : downloading ? 'Downloading scanner…' : 'Preparing scanner…'}
        </span>
      </div>
      <div className="h-2 bg-slate-200 rounded-full overflow-hidden">
        {showBar
          ? <div className="h-full bg-navy rounded-full transition-all" style={{ width: `${pct}%` }} />
          : <div className="h-full w-1/3 bg-navy/70 rounded-full animate-pulse" />}
      </div>
      <p className="text-[11px] text-slate-400">
        First-time setup downloads ~9&nbsp;MB and is cached for next time.
        {showBar && ` ${(prog.loaded / 1e6).toFixed(1)} MB`}
      </p>
      <button type="button" onClick={onSkip}
        className="text-[11px] text-navy font-semibold hover:underline">
        Skip and scan without auto-crop
      </button>
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
