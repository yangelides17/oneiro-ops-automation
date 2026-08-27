import { useState, useRef, useEffect } from 'react'
import { Document, Page, pdfjs } from 'react-pdf'

pdfjs.GlobalWorkerOptions.workerSrc = `//unpkg.com/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`

/**
 * PDF first-page thumbnail. Renders page 1 scaled to fill the
 * container width, clipped to a fixed height. Like Google Drive.
 */
export default function PdfThumbnail({ url, height = 160, className = '' }) {
  const [visible, setVisible] = useState(false)
  const [loaded, setLoaded] = useState(false)
  const [error, setError] = useState(false)
  const [containerWidth, setContainerWidth] = useState(0)
  const ref = useRef(null)

  // Lazy load
  useEffect(() => {
    if (!ref.current) return
    setContainerWidth(ref.current.offsetWidth)
    const observer = new IntersectionObserver(
      ([entry]) => {
        if (entry.isIntersecting) {
          setVisible(true)
          setContainerWidth(entry.target.offsetWidth)
          observer.disconnect()
        }
      },
      { rootMargin: '200px' }
    )
    observer.observe(ref.current)
    return () => observer.disconnect()
  }, [])

  return (
    <div
      ref={ref}
      className={`overflow-hidden bg-white relative ${className}`}
      style={{ height, maxHeight: height }}
    >
      {(!visible || (!loaded && !error)) && (
        <div className="absolute inset-0 flex items-center justify-center bg-slate-50">
          {!visible
            ? <span className="text-2xl text-slate-200">📄</span>
            : <div className="w-4 h-4 border-2 border-slate-200 border-t-slate-400 rounded-full animate-spin" />
          }
        </div>
      )}

      {error && (
        <div className="absolute inset-0 flex items-center justify-center bg-slate-50">
          <span className="text-2xl text-slate-200">📄</span>
        </div>
      )}

      {visible && !error && containerWidth > 0 && (
        <Document
          file={url}
          onLoadError={() => setError(true)}
          loading=""
        >
          <Page
            pageNumber={1}
            width={containerWidth}
            renderTextLayer={false}
            renderAnnotationLayer={false}
            onRenderSuccess={() => setLoaded(true)}
          />
        </Document>
      )}
    </div>
  )
}
