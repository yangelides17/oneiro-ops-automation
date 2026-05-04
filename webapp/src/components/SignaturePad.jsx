import { useRef, useState, useEffect, useCallback } from 'react'

/**
 * SignaturePad — touchscreen / mouse-friendly canvas signature capture.
 *
 * Inline mode is a small in-form canvas. The "Expand" button opens a
 * fullscreen modal with a much larger canvas; existing strokes carry
 * over both directions. The modal is helpful on mobile (rotate the
 * phone, use the whole screen for a real signature) and on desktop
 * (just more room).
 *
 * Stroke style is true black (#000000) at 1.5px to read like pen on
 * paper rather than a thick highlighter.
 *
 * Props:
 *   label     — field label text
 *   onChange  — called with base64 PNG data URL whenever drawing ends,
 *               or null when the pad is cleared.
 */

const STROKE_COLOR = '#000000'
const STROKE_WIDTH = 1.5

// Set up a canvas's drawing context: device-pixel-ratio scaling so
// strokes stay sharp on retina, plus our pen style.
function configureCanvas(canvas) {
  const dpr  = window.devicePixelRatio || 1
  const rect = canvas.getBoundingClientRect()
  canvas.width  = rect.width  * dpr
  canvas.height = rect.height * dpr
  const ctx = canvas.getContext('2d')
  ctx.setTransform(1, 0, 0, 1, 0, 0)
  ctx.scale(dpr, dpr)
  ctx.strokeStyle = STROKE_COLOR
  ctx.lineWidth   = STROKE_WIDTH
  ctx.lineCap     = 'round'
  ctx.lineJoin    = 'round'
  return ctx
}

// Draw a base64 PNG onto the given canvas, scaled to fill it. Used
// when transferring strokes between inline ↔ expanded canvases and
// when the expanded canvas resizes (e.g. phone rotation).
function paintDataUrl(canvas, dataUrl, onDone) {
  if (!dataUrl) {
    onDone?.()
    return
  }
  const img = new Image()
  img.onload = () => {
    const ctx  = canvas.getContext('2d')
    const rect = canvas.getBoundingClientRect()
    ctx.drawImage(img, 0, 0, rect.width, rect.height)
    onDone?.()
  }
  img.onerror = () => onDone?.()
  img.src = dataUrl
}

function clearCanvas(canvas) {
  const dpr = window.devicePixelRatio || 1
  canvas.getContext('2d').clearRect(0, 0, canvas.width / dpr, canvas.height / dpr)
}

// True if the canvas has any non-zero alpha pixels (i.e. ink).
function canvasHasInk(canvas) {
  if (!canvas) return false
  const ctx = canvas.getContext('2d')
  const { width, height } = canvas
  if (!width || !height) return false
  try {
    const data = ctx.getImageData(0, 0, width, height).data
    for (let i = 3; i < data.length; i += 4) {
      if (data[i] !== 0) return true
    }
    return false
  } catch {
    return false
  }
}

// Pointer position relative to the canvas (mouse and touch).
function pointFromEvent(canvas, e) {
  const rect = canvas.getBoundingClientRect()
  const src  = e.touches ? e.touches[0] : e
  return { x: src.clientX - rect.left, y: src.clientY - rect.top }
}

export default function SignaturePad({ label, onChange }) {
  const inlineRef = useRef(null)
  const drawing   = useRef(false)
  const lastPt    = useRef(null)
  const [signed,   setSigned]   = useState(false)
  const [expanded, setExpanded] = useState(false)
  // Snapshot of the inline canvas at expand time, so the modal can
  // pre-render existing strokes.
  const expandSeed = useRef(null)

  // Initial inline-canvas setup
  useEffect(() => {
    if (inlineRef.current) configureCanvas(inlineRef.current)
  }, [])

  // ── Inline-canvas drawing handlers ───────────────────────────
  const onPointerDown = (e) => {
    e.preventDefault()
    drawing.current = true
    lastPt.current  = pointFromEvent(inlineRef.current, e)
  }
  const onPointerMove = (e) => {
    if (!drawing.current) return
    e.preventDefault()
    const pt  = pointFromEvent(inlineRef.current, e)
    const ctx = inlineRef.current.getContext('2d')
    ctx.beginPath()
    ctx.moveTo(lastPt.current.x, lastPt.current.y)
    ctx.lineTo(pt.x, pt.y)
    ctx.stroke()
    lastPt.current = pt
    if (!signed) setSigned(true)
  }
  const onPointerEnd = () => {
    if (!drawing.current) return
    drawing.current = false
    onChange?.(inlineRef.current.toDataURL('image/png'))
  }

  const clearInline = (e) => {
    e?.stopPropagation()
    clearCanvas(inlineRef.current)
    setSigned(false)
    onChange?.(null)
  }

  // Open the fullscreen modal — capture the current inline drawing as
  // a seed so the modal can pre-render it.
  const openExpand = () => {
    expandSeed.current = signed ? inlineRef.current.toDataURL('image/png') : null
    setExpanded(true)
  }

  // Modal handed back: paint result onto inline canvas + propagate.
  // resultDataUrl semantics:
  //   string   → user accepted strokes; transfer + onChange
  //   null     → user cleared in modal; clear inline + onChange(null)
  //   undefined→ user cancelled; leave inline untouched
  const closeExpand = (resultDataUrl) => {
    setExpanded(false)
    if (resultDataUrl === undefined) return
    if (resultDataUrl === null) {
      clearCanvas(inlineRef.current)
      setSigned(false)
      onChange?.(null)
      return
    }
    const canvas = inlineRef.current
    clearCanvas(canvas)
    paintDataUrl(canvas, resultDataUrl, () => {
      setSigned(true)
      onChange?.(canvas.toDataURL('image/png'))
    })
  }

  return (
    <>
      <div className="space-y-1">
        <div className="flex justify-between items-center">
          <label className="field-label">{label}</label>
          <div className="flex items-center gap-3">
            {signed && (
              <button
                type="button"
                onClick={clearInline}
                className="text-[11px] text-red-400 hover:text-red-600 transition-colors font-medium"
              >
                Clear
              </button>
            )}
            <button
              type="button"
              onClick={openExpand}
              aria-label="Expand signature pad"
              className="text-slate-400 hover:text-navy transition-colors p-1 -m-1"
            >
              <ExpandIcon />
            </button>
          </div>
        </div>

        <div
          className={`relative border-2 rounded-xl overflow-hidden bg-white transition-colors
            ${signed ? 'border-navy/30' : 'border-dashed border-slate-200'}`}
        >
          <canvas
            ref={inlineRef}
            style={{ width: '100%', height: '100px', display: 'block', touchAction: 'none' }}
            onMouseDown={onPointerDown}
            onMouseMove={onPointerMove}
            onMouseUp={onPointerEnd}
            onMouseLeave={onPointerEnd}
            onTouchStart={onPointerDown}
            onTouchMove={onPointerMove}
            onTouchEnd={onPointerEnd}
          />
          {!signed && (
            <div className="absolute inset-0 flex items-center justify-center pointer-events-none">
              <span className="text-slate-300 text-sm select-none">Sign here</span>
            </div>
          )}
        </div>

        {signed && (
          <p className="text-[10px] text-green-600 font-semibold flex items-center gap-1">
            <span>✓</span> Signature captured
          </p>
        )}
      </div>

      {expanded && (
        <ExpandedSignatureModal
          label={label}
          seedDataUrl={expandSeed.current}
          onClose={closeExpand}
        />
      )}
    </>
  )
}

// ── Fullscreen expand modal ────────────────────────────────────────
//
// Independent canvas + drawing handlers (no shared state with the
// inline pad). On open we paint `seedDataUrl` so the user picks up
// where they left off; on Done we hand the resulting PNG back via
// `onClose(dataUrl)`. On Cancel `onClose(undefined)` — the inline
// pad is left untouched. Tapping Clear in the modal then Done returns
// null, which clears the inline pad.
function ExpandedSignatureModal({ label, seedDataUrl, onClose }) {
  const canvasRef = useRef(null)
  const drawing   = useRef(false)
  const lastPt    = useRef(null)
  const [hasInk,  setHasInk]  = useState(!!seedDataUrl)
  // Latest data URL captured after every stroke (and after resize)
  // so phone rotation can re-paint without losing strokes.
  const latestRef = useRef(seedDataUrl)

  const setupAndPaint = useCallback((paintWith) => {
    const canvas = canvasRef.current
    if (!canvas) return
    configureCanvas(canvas)
    if (paintWith) {
      paintDataUrl(canvas, paintWith, () => {
        latestRef.current = canvas.toDataURL('image/png')
      })
    }
  }, [])

  // Initial setup + paint the seed.
  useEffect(() => {
    setupAndPaint(seedDataUrl)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // Resize / orientation changes — preserve current strokes by
  // capturing → resizing → repainting.
  useEffect(() => {
    const onResize = () => {
      const canvas = canvasRef.current
      if (!canvas) return
      const snapshot = latestRef.current
      setupAndPaint(snapshot)
    }
    window.addEventListener('resize', onResize)
    window.addEventListener('orientationchange', onResize)
    return () => {
      window.removeEventListener('resize', onResize)
      window.removeEventListener('orientationchange', onResize)
    }
  }, [setupAndPaint])

  const onPointerDown = (e) => {
    e.preventDefault()
    drawing.current = true
    lastPt.current  = pointFromEvent(canvasRef.current, e)
  }
  const onPointerMove = (e) => {
    if (!drawing.current) return
    e.preventDefault()
    const pt  = pointFromEvent(canvasRef.current, e)
    const ctx = canvasRef.current.getContext('2d')
    ctx.beginPath()
    ctx.moveTo(lastPt.current.x, lastPt.current.y)
    ctx.lineTo(pt.x, pt.y)
    ctx.stroke()
    lastPt.current = pt
    if (!hasInk) setHasInk(true)
  }
  const onPointerEnd = () => {
    if (!drawing.current) return
    drawing.current = false
    latestRef.current = canvasRef.current.toDataURL('image/png')
  }

  const clearAll = () => {
    clearCanvas(canvasRef.current)
    setHasInk(false)
    latestRef.current = null
  }

  const handleDone = () => {
    if (!canvasRef.current) return onClose(undefined)
    const ink = canvasHasInk(canvasRef.current)
    onClose(ink ? canvasRef.current.toDataURL('image/png') : null)
  }

  return (
    <div className="fixed inset-0 z-[100] bg-white flex flex-col">
      {/* Top bar: Cancel | label | Done */}
      <div className="flex items-center justify-between px-4 py-3 border-b border-slate-200">
        <button
          type="button"
          onClick={() => onClose(undefined)}
          className="text-sm text-slate-600 font-medium px-2 py-1"
        >
          Cancel
        </button>
        <span className="text-xs font-bold uppercase tracking-wider text-navy text-center
                         flex-1 mx-3 truncate">
          {label}
        </span>
        <button
          type="button"
          onClick={handleDone}
          className="text-sm text-navy font-bold px-2 py-1"
        >
          Done
        </button>
      </div>

      {/* Canvas fills available space. touchAction:none stops iOS from
          interpreting drags as scroll/zoom gestures. */}
      <div className="relative flex-1 bg-white">
        <canvas
          ref={canvasRef}
          style={{ width: '100%', height: '100%', display: 'block', touchAction: 'none' }}
          onMouseDown={onPointerDown}
          onMouseMove={onPointerMove}
          onMouseUp={onPointerEnd}
          onMouseLeave={onPointerEnd}
          onTouchStart={onPointerDown}
          onTouchMove={onPointerMove}
          onTouchEnd={onPointerEnd}
        />
        {!hasInk && (
          <div className="absolute inset-0 flex items-center justify-center pointer-events-none">
            <span className="text-slate-300 text-base select-none">
              Sign anywhere — rotate phone for more room
            </span>
          </div>
        )}
      </div>

      {/* Footer Clear button — visible whenever there's ink. */}
      <div className="flex items-center justify-center px-4 py-3 border-t border-slate-200 min-h-[52px]">
        {hasInk && (
          <button
            type="button"
            onClick={clearAll}
            className="text-sm text-red-500 hover:text-red-700 font-semibold px-4 py-1"
          >
            Clear
          </button>
        )}
      </div>
    </div>
  )
}

function ExpandIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none"
         stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="15 3 21 3 21 9" />
      <polyline points="9 21 3 21 3 15" />
      <line x1="21" y1="3" x2="14" y2="10" />
      <line x1="3" y1="21" x2="10" y2="14" />
    </svg>
  )
}
