import { useRef, useState, useEffect } from 'react'

/**
 * SignaturePad — touchscreen/mouse-friendly canvas signature capture.
 *
 * Props:
 *   label     — field label text
 *   onChange  — called with base64 PNG data URL whenever drawing ends,
 *               or null when cleared
 */
export default function SignaturePad({ label, onChange }) {
  const canvasRef  = useRef(null)
  const drawing    = useRef(false)
  const lastPt     = useRef(null)
  const [signed, setSigned] = useState(false)

  // Scale canvas for device pixel ratio so it looks sharp on retina
  useEffect(() => {
    const canvas = canvasRef.current
    const dpr    = window.devicePixelRatio || 1
    const rect   = canvas.getBoundingClientRect()
    canvas.width  = rect.width  * dpr
    canvas.height = rect.height * dpr
    const ctx = canvas.getContext('2d')
    ctx.scale(dpr, dpr)
    ctx.strokeStyle = '#1B2A4A'
    ctx.lineWidth   = 2.5
    ctx.lineCap     = 'round'
    ctx.lineJoin    = 'round'
  }, [])

  const getPoint = (e) => {
    const canvas = canvasRef.current
    const rect   = canvas.getBoundingClientRect()
    const src    = e.touches ? e.touches[0] : e
    return {
      x: src.clientX - rect.left,
      y: src.clientY - rect.top
    }
  }

  const startDraw = (e) => {
    e.preventDefault()
    drawing.current = true
    lastPt.current  = getPoint(e)
  }

  const draw = (e) => {
    e.preventDefault()
    if (!drawing.current) return
    const pt  = getPoint(e)
    const ctx = canvasRef.current.getContext('2d')
    ctx.beginPath()
    ctx.moveTo(lastPt.current.x, lastPt.current.y)
    ctx.lineTo(pt.x, pt.y)
    ctx.stroke()
    lastPt.current = pt
    if (!signed) setSigned(true)
  }

  const endDraw = (e) => {
    if (!drawing.current) return
    drawing.current = false
    // Notify parent with PNG data URL
    onChange?.(canvasRef.current.toDataURL('image/png'))
  }

  const clear = (e) => {
    e.stopPropagation()
    const canvas = canvasRef.current
    const dpr    = window.devicePixelRatio || 1
    canvas.getContext('2d').clearRect(0, 0, canvas.width / dpr, canvas.height / dpr)
    setSigned(false)
    onChange?.(null)
  }

  return (
    <div className="space-y-1">
      <div className="flex justify-between items-center">
        <label className="field-label">{label}</label>
        {signed && (
          <button
            type="button"
            onClick={clear}
            className="text-[11px] text-red-400 hover:text-red-600 transition-colors font-medium"
          >
            Clear
          </button>
        )}
      </div>

      <div
        className={`relative border-2 rounded-xl overflow-hidden bg-white transition-colors
          ${signed ? 'border-navy/30' : 'border-dashed border-slate-200'}`}
      >
        <canvas
          ref={canvasRef}
          style={{ width: '100%', height: '100px', display: 'block', touchAction: 'none' }}
          onMouseDown={startDraw}
          onMouseMove={draw}
          onMouseUp={endDraw}
          onMouseLeave={endDraw}
          onTouchStart={startDraw}
          onTouchMove={draw}
          onTouchEnd={endDraw}
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
  )
}
