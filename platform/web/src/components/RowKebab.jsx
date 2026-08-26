import { useEffect, useRef, useState } from 'react'

/**
 * RowKebab — a vertical-three-dots (⋮) button that opens a small menu.
 *
 * Props:
 *   items — array of { label, onClick, danger? }. Each click closes the menu.
 *   disabled — bool; when true, the button is inert.
 */
export default function RowKebab({ items = [], disabled = false }) {
  const [open, setOpen] = useState(false)
  const wrapRef = useRef(null)

  useEffect(() => {
    if (!open) return
    const handler = (e) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) setOpen(false)
    }
    const esc = (e) => { if (e.key === 'Escape') setOpen(false) }
    document.addEventListener('mousedown', handler)
    document.addEventListener('keydown', esc)
    return () => {
      document.removeEventListener('mousedown', handler)
      document.removeEventListener('keydown', esc)
    }
  }, [open])

  return (
    <div ref={wrapRef} className="relative">
      <button
        type="button"
        disabled={disabled}
        onClick={() => setOpen(v => !v)}
        className="w-7 h-7 flex items-center justify-center rounded-md
                   text-slate-400 hover:text-slate-700 hover:bg-slate-100
                   disabled:opacity-30 disabled:cursor-not-allowed
                   transition-colors"
        aria-label="Row actions"
      >
        {/* ⋮ glyph — custom SVG so it renders cleanly at any zoom */}
        <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
          <circle cx="8" cy="3" r="1.5" />
          <circle cx="8" cy="8" r="1.5" />
          <circle cx="8" cy="13" r="1.5" />
        </svg>
      </button>

      {open && (
        <div className="absolute right-0 mt-1 z-30 min-w-[140px]
                        bg-white rounded-lg shadow-lg border border-slate-200
                        py-1 text-sm overflow-hidden">
          {items.map((item, i) => (
            <button
              key={i}
              type="button"
              onClick={() => { setOpen(false); item.onClick?.() }}
              className={`w-full text-left px-3 py-2 transition-colors
                ${item.danger
                  ? 'text-red-600 hover:bg-red-50'
                  : 'text-slate-700 hover:bg-slate-50'}`}
            >
              {item.label}
            </button>
          ))}
        </div>
      )}
    </div>
  )
}
