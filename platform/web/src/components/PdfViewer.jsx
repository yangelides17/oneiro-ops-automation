import { useState, useCallback, useEffect, useRef } from 'react'
import { Document, Page, pdfjs } from 'react-pdf'
import 'react-pdf/dist/esm/Page/AnnotationLayer.css'
import 'react-pdf/dist/esm/Page/TextLayer.css'

// Configure pdf.js worker
pdfjs.GlobalWorkerOptions.workerSrc = `//unpkg.com/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`

/**
 * Inline PDF viewer using react-pdf (Mozilla pdf.js).
 *
 * Renders PDFs with a white background and custom controls,
 * matching the app's design language. No Chrome dark-mode issues.
 *
 * Props:
 *   url       — same-origin proxy URL for the PDF
 *   filename  — display name in the header bar
 *   collapsed — whether the panel is collapsed (controlled)
 *   onToggle  — callback to toggle collapsed state
 *   onClose   — callback to close/remove the viewer (optional)
 *   onRefresh — async callback for reload (optional)
 *   className — additional Tailwind classes (optional)
 */
export default function PdfViewer({
  url,
  filename,
  collapsed = false,
  onToggle,
  onClose,
  onRefresh,
  fillParent = false,
  className = '',
}) {
  const [numPages, setNumPages] = useState(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState(false)
  const [containerWidth, setContainerWidth] = useState(null)
  const containerRef = useRef(null)

  // Track container width for responsive page sizing
  useEffect(() => {
    if (!containerRef.current || collapsed) return
    const observer = new ResizeObserver(entries => {
      for (const entry of entries) {
        setContainerWidth(entry.contentRect.width)
      }
    })
    observer.observe(containerRef.current)
    return () => observer.disconnect()
  }, [collapsed])

  // Reset when URL changes
  useEffect(() => {
    setLoading(true)
    setError(false)
    setNumPages(null)
  }, [url])

  const onDocumentLoadSuccess = useCallback(({ numPages: n }) => {
    setNumPages(n)
    setLoading(false)
    setError(false)
  }, [])

  const onDocumentLoadError = useCallback(() => {
    setLoading(false)
    setError(true)
  }, [])

  const handleRefresh = useCallback(async () => {
    setLoading(true)
    setError(false)
    setNumPages(null)
    if (onRefresh) {
      try { await onRefresh() } catch { /* ignore */ }
    }
    // Force re-render by toggling a key (handled by parent re-passing url)
  }, [onRefresh])

  if (!url) return null

  return (
    <div className={`border border-slate-200 rounded-xl bg-white overflow-hidden flex flex-col h-full ${className}`}>
      {/* Header bar */}
      <div className="flex items-center justify-between gap-2 px-3 py-2 bg-slate-50 border-b border-slate-200">
        <button
          type="button"
          onClick={onToggle}
          className="flex items-center gap-2 min-w-0 flex-1 text-left"
        >
          <span className="text-base flex-shrink-0">📄</span>
          <span className="text-sm font-semibold text-slate-700 truncate">
            {filename || 'Document'}
          </span>
          <span className="text-xs text-slate-400 flex-shrink-0">
            {collapsed ? '▸' : '▾'}
          </span>
        </button>

        <div className="flex items-center gap-1 flex-shrink-0">
          {/* Page count indicator */}
          {!collapsed && numPages && (
            <span className="text-xs text-slate-400 tabular-nums mr-2">
              {numPages} page{numPages === 1 ? '' : 's'}
            </span>
          )}

          {error && (
            <button
              type="button"
              onClick={handleRefresh}
              className="text-xs font-semibold px-2 py-1 rounded-lg
                         text-amber-700 bg-amber-50 border border-amber-200
                         hover:bg-amber-100"
            >
              Retry
            </button>
          )}
          {!collapsed && (
            <a
              href={url}
              target="_blank"
              rel="noopener noreferrer"
              className="text-xs font-semibold px-2 py-1 rounded-lg
                         text-slate-600 hover:bg-slate-100"
              title="Open in new tab"
            >
              ↗
            </a>
          )}
          {onClose && (
            <button
              type="button"
              onClick={onClose}
              className="w-6 h-6 flex items-center justify-center rounded-full
                         text-slate-400 hover:text-red-500 hover:bg-red-50"
              title="Close"
            >
              ×
            </button>
          )}
        </div>
      </div>

      {/* PDF content */}
      {!collapsed && (
        <div
          ref={containerRef}
          className="relative overflow-auto bg-slate-100 flex-1"
          style={{ minHeight: 200, ...(fillParent ? {} : { maxHeight: '60vh' }) }}
        >
          {loading && (
            <div className="absolute inset-0 flex items-center justify-center bg-slate-50/80 z-10">
              <div className="flex flex-col items-center gap-2">
                <div className="w-6 h-6 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
                <span className="text-xs text-slate-400">Loading PDF…</span>
              </div>
            </div>
          )}

          {error && (
            <div className="flex flex-col items-center justify-center py-8 gap-2">
              <span className="text-2xl">⚠</span>
              <p className="text-sm text-slate-500">Failed to load PDF</p>
              <button
                type="button"
                onClick={handleRefresh}
                className="text-xs font-bold px-3 py-1.5 rounded-lg
                           bg-navy text-white hover:opacity-90"
              >
                Retry
              </button>
            </div>
          )}

          {!error && (
            <Document
              file={url}
              onLoadSuccess={onDocumentLoadSuccess}
              onLoadError={onDocumentLoadError}
              loading=""
              className="flex flex-col items-center py-3 gap-3"
            >
              {numPages && Array.from({ length: numPages }, (_, i) => (
                <Page
                  key={i + 1}
                  pageNumber={i + 1}
                  width={containerWidth ? Math.min(containerWidth - 24, 900) : undefined}
                  renderTextLayer={false}
                  renderAnnotationLayer={false}
                  className="shadow-md"
                />
              ))}
            </Document>
          )}
        </div>
      )}
    </div>
  )
}
