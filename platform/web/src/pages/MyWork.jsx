import { useState, useEffect, useMemo, useCallback } from 'react'
import { Link } from 'react-router-dom'
import { GoogleMap, useJsApiLoader, MarkerF, InfoWindowF } from '@react-google-maps/api'
import StatusBadge from '../components/StatusBadge'
import PdfViewer from '../components/PdfViewer'

const NYC_CENTER = { lat: 40.7128, lng: -74.0060 }
const MAP_OPTIONS = {
  disableDefaultUI: true,
  zoomControl: true,
  mapTypeControl: false,
  streetViewControl: false,
  fullscreenControl: false,
}

const STATUS_ORDER = { received: 0, dispatched: 1, in_progress: 2, completed: 3 }

const PIN_COLORS = {
  received:    '#3b82f6',
  dispatched:  '#f59e0b',
  in_progress: '#f97316',
  completed:   '#22c55e',
}

function buildPin(status) {
  const s = String(status || '').toLowerCase().replace(/\s+/g, '_')
  const color = PIN_COLORS[s] || '#3b82f6'
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="28" height="40" viewBox="0 0 24 36">
    <path d="M12 0C5.4 0 0 5.4 0 12c0 9 12 24 12 24s12-15 12-24C24 5.4 18.6 0 12 0z" fill="${color}" stroke="#fff" stroke-width="1.5"/>
    <circle cx="12" cy="12" r="5" fill="#fff" opacity="0.9"/>
  </svg>`
  return {
    url: 'data:image/svg+xml;charset=UTF-8,' + encodeURIComponent(svg),
    scaledSize: { width: 28, height: 40 },
    anchor: { x: 14, y: 40 },
  }
}

function getDirectionsUrl(wo) {
  if (wo.latitude && wo.longitude) {
    return `https://www.google.com/maps/dir/?api=1&destination=${wo.latitude},${wo.longitude}`
  }
  if (wo.location) {
    return `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(wo.location)}`
  }
  return null
}

export default function MyWork() {
  const apiKey = import.meta.env.VITE_GOOGLE_MAPS_BROWSER_KEY
  const { isLoaded } = useJsApiLoader({ id: 'google-maps-script', googleMapsApiKey: apiKey || '' })

  const [wos, setWos] = useState([])
  const [loading, setLoading] = useState(true)
  const [mobileView, setMobileView] = useState('list') // 'list' | 'map'
  const [expandedId, setExpandedId] = useState(null)
  const [activePin, setActivePin] = useState(null)
  const [pdfWo, setPdfWo] = useState(null)
  const [pdfUrl, setPdfUrl] = useState(null)

  const loadWork = useCallback(async () => {
    setLoading(true)
    try {
      const res = await fetch('/api/wos/my-work')
      const data = await res.json()
      setWos(data.wos || [])
    } catch (e) {
      console.error('Failed to load my work:', e)
    } finally {
      setLoading(false)
    }
  }, [])

  useEffect(() => { loadWork() }, [loadWork])

  // Sort: active first, completed last
  const sortedWos = useMemo(() => {
    return [...wos].sort((a, b) => {
      const sa = String(a.status || '').toLowerCase().replace(/\s+/g, '_')
      const sb = String(b.status || '').toLowerCase().replace(/\s+/g, '_')
      return (STATUS_ORDER[sa] ?? 1) - (STATUS_ORDER[sb] ?? 1)
    })
  }, [wos])

  const mappedWos = useMemo(() => wos.filter(w => w.latitude && w.longitude), [wos])

  const statusCounts = useMemo(() => {
    const c = { total: wos.length, active: 0, completed: 0 }
    for (const w of wos) {
      if (String(w.status || '').toLowerCase() === 'completed') c.completed++
      else c.active++
    }
    return c
  }, [wos])

  const openPdf = useCallback((wo) => {
    setPdfWo(wo)
    setPdfUrl(null)
    if (wo.scanFileKey) {
      fetch(`/api/wos/${encodeURIComponent(wo.id)}/files`)
        .then(r => r.json())
        .then(d => setPdfUrl(d.scan?.url || null))
        .catch(() => setPdfUrl(null))
    }
  }, [])

  const closePdf = () => { setPdfWo(null); setPdfUrl(null) }

  const toggleExpand = (woId) => {
    setExpandedId(prev => prev === woId ? null : woId)
  }

  const mapCenter = useMemo(() => {
    if (mappedWos.length > 0) return { lat: Number(mappedWos[0].latitude), lng: Number(mappedWos[0].longitude) }
    return NYC_CENTER
  }, [mappedWos])

  return (
    <div className="h-[calc(100vh-56px)] flex flex-col">
      {/* ── Header ─────────────────────────────────────────────── */}
      <div className="flex items-center justify-between gap-2 px-4 py-2.5 border-b border-slate-200 bg-white flex-shrink-0">
        <div className="min-w-0">
          <h1 className="text-lg font-black text-navy leading-tight">My Work</h1>
          <p className="text-[11px] text-slate-400 leading-tight mt-0.5">
            {statusCounts.active > 0 && <><span className="font-bold text-slate-600">{statusCounts.active}</span> active</>}
            {statusCounts.completed > 0 && <>{statusCounts.active > 0 && ' · '}<span className="text-green-600">{statusCounts.completed} done</span></>}
            {statusCounts.total === 0 && 'No work orders'}
          </p>
        </div>
        <div className="flex items-center gap-2 flex-shrink-0">
          {/* Mobile view toggle */}
          <div className="sm:hidden flex bg-slate-100 rounded-lg p-0.5">
            <button
              onClick={() => setMobileView('list')}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-all ${
                mobileView === 'list' ? 'bg-white text-navy shadow-sm' : 'text-slate-500'
              }`}
            >
              List
            </button>
            <button
              onClick={() => setMobileView('map')}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-all ${
                mobileView === 'map' ? 'bg-white text-navy shadow-sm' : 'text-slate-500'
              }`}
            >
              Map
            </button>
          </div>
          <button onClick={loadWork} disabled={loading}
                  className="p-2 rounded-lg hover:bg-slate-100 active:bg-slate-200 transition-colors"
                  title="Refresh">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"
                 className={`text-slate-500 ${loading ? 'animate-spin' : ''}`}>
              <path d="M21 12a9 9 0 11-6.219-8.56" /><path d="M21 3v6h-6" />
            </svg>
          </button>
        </div>
      </div>

      {/* ── Main content ───────────────────────────────────────── */}
      <div className="flex-1 flex flex-col sm:flex-row min-h-0">
        {/* Map — full screen on mobile (map view), left half on desktop */}
        <div className={`
          ${mobileView === 'map' ? 'flex-1' : 'hidden'}
          sm:block sm:w-1/2 sm:flex-none min-h-0
        `}>
          {isLoaded ? (
            <GoogleMap
              mapContainerStyle={{ width: '100%', height: '100%' }}
              center={mapCenter}
              zoom={mappedWos.length > 0 ? 13 : 11}
              options={MAP_OPTIONS}
            >
              {mappedWos.map(w => (
                <MarkerF
                  key={w.id}
                  position={{ lat: Number(w.latitude), lng: Number(w.longitude) }}
                  icon={buildPin(w.status)}
                  onClick={() => setActivePin(w)}
                />
              ))}
              {activePin && activePin.latitude && (
                <InfoWindowF
                  position={{ lat: Number(activePin.latitude), lng: Number(activePin.longitude) }}
                  onCloseClick={() => setActivePin(null)}
                >
                  <MapPopover
                    wo={activePin}
                    onViewPdf={() => { openPdf(activePin); setActivePin(null) }}
                  />
                </InfoWindowF>
              )}
            </GoogleMap>
          ) : (
            <div className="flex items-center justify-center h-full bg-slate-50">
              <div className="w-8 h-8 border-[3px] border-slate-200 border-t-navy rounded-full animate-spin" />
            </div>
          )}
        </div>

        {/* WO list — full screen on mobile (list view), right half on desktop */}
        <div className={`
          ${mobileView === 'list' ? 'flex-1' : 'hidden'}
          sm:block sm:w-1/2 sm:flex-none sm:border-l border-slate-200
          overflow-y-auto bg-slate-50 sm:bg-white min-h-0
        `}>
          {loading && (
            <div className="flex items-center justify-center py-12">
              <div className="w-7 h-7 border-[3px] border-slate-200 border-t-navy rounded-full animate-spin" />
            </div>
          )}

          {!loading && wos.length === 0 && (
            <div className="flex flex-col items-center justify-center py-16 px-6 text-center">
              <div className="w-16 h-16 bg-slate-100 rounded-2xl flex items-center justify-center mb-4">
                <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="1.5" strokeLinecap="round">
                  <path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2" />
                  <rect x="9" y="3" width="6" height="4" rx="1" />
                </svg>
              </div>
              <p className="text-sm font-semibold text-slate-600">No work orders assigned</p>
              <p className="text-xs text-slate-400 mt-1 max-w-[240px]">
                Your field manager will assign work orders to you from the assignment board.
              </p>
            </div>
          )}

          {!loading && sortedWos.length > 0 && (
            <div className="p-2 sm:p-0 space-y-2 sm:space-y-0">
              {sortedWos.map(wo => (
                <WOCard
                  key={wo.id}
                  wo={wo}
                  expanded={expandedId === wo.id}
                  onToggle={() => toggleExpand(wo.id)}
                  onViewPdf={() => openPdf(wo)}
                />
              ))}
            </div>
          )}
        </div>
      </div>

      {/* ── PDF overlay ────────────────────────────────────────── */}
      {pdfWo && (
        <>
          {/* Mobile: full-screen overlay */}
          <div className="fixed inset-0 z-50 bg-white flex flex-col sm:hidden">
            <div className="flex items-center gap-3 px-4 py-3 border-b border-slate-200 flex-shrink-0">
              <button onClick={closePdf}
                      className="p-1.5 -ml-1.5 rounded-lg hover:bg-slate-100 active:bg-slate-200">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round">
                  <path d="M19 12H5" /><path d="M12 19l-7-7 7-7" />
                </svg>
              </button>
              <div className="min-w-0 flex-1">
                <p className="text-sm font-bold text-navy truncate">{pdfWo.woNumber} Scan</p>
                <p className="text-[10px] text-slate-400 truncate">{pdfWo.location || ''}</p>
              </div>
            </div>
            <div className="flex-1 min-h-0">
              {pdfUrl ? (
                <PdfViewer url={pdfUrl} filename={`${pdfWo.woNumber} Scan`}
                           collapsed={false} onToggle={() => {}} onClose={closePdf} fillParent />
              ) : pdfWo.scanFileKey ? (
                <div className="flex items-center justify-center h-full">
                  <div className="w-7 h-7 border-[3px] border-slate-200 border-t-navy rounded-full animate-spin" />
                </div>
              ) : (
                <div className="flex items-center justify-center h-full text-sm text-slate-400">
                  No scan available for this work order
                </div>
              )}
            </div>
          </div>

          {/* Desktop: slide-in panel from right */}
          <div className="hidden sm:flex fixed top-14 right-0 bottom-0 w-1/2 z-40 bg-white border-l border-slate-200 shadow-xl flex-col">
            <div className="flex items-center gap-3 px-4 py-2 border-b border-slate-200 flex-shrink-0">
              <button onClick={closePdf} className="p-1 -ml-1 rounded-lg hover:bg-slate-100">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round">
                  <line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" />
                </svg>
              </button>
              <span className="text-sm font-bold text-navy">{pdfWo.woNumber} Scan</span>
            </div>
            <div className="flex-1 min-h-0">
              {pdfUrl ? (
                <PdfViewer url={pdfUrl} filename={`${pdfWo.woNumber} Scan`}
                           collapsed={false} onToggle={() => {}} onClose={closePdf} fillParent />
              ) : pdfWo.scanFileKey ? (
                <div className="flex items-center justify-center h-full">
                  <div className="w-7 h-7 border-[3px] border-slate-200 border-t-navy rounded-full animate-spin" />
                </div>
              ) : (
                <div className="flex items-center justify-center h-full text-sm text-slate-400">
                  No scan available
                </div>
              )}
            </div>
          </div>
        </>
      )}
    </div>
  )
}

// ── WO Card ───────────────────────────────────────────────────

function WOCard({ wo, expanded, onToggle, onViewPdf }) {
  const isCompleted = String(wo.status || '').toLowerCase() === 'completed'
  const streets = [wo.fromStreet, wo.toStreet].filter(Boolean).join(' \u2192 ')
  const directionsUrl = getDirectionsUrl(wo)

  return (
    <div className={`
      sm:border-b sm:border-slate-100
      rounded-xl sm:rounded-none bg-white sm:bg-transparent
      ${isCompleted ? 'opacity-60' : ''}
      ${expanded ? 'ring-2 ring-navy/20 sm:ring-0 sm:bg-navy/5' : ''}
      transition-all
    `}>
      {/* Main row */}
      <button
        type="button"
        onClick={onToggle}
        className="w-full text-left px-4 py-3.5 sm:py-3 flex items-start gap-3 active:bg-slate-50 transition-colors"
      >
        {/* Status dot */}
        <div className="mt-1.5 flex-shrink-0">
          <div className={`w-2.5 h-2.5 rounded-full ${
            isCompleted ? 'bg-green-400' :
            String(wo.status || '').toLowerCase().replace(/\s+/g, '_') === 'in_progress' ? 'bg-orange-400' :
            String(wo.status || '').toLowerCase() === 'dispatched' ? 'bg-amber-400' :
            'bg-blue-400'
          }`} />
        </div>

        {/* Content */}
        <div className="flex-1 min-w-0">
          <div className="flex items-center justify-between gap-2">
            <span className="text-sm font-bold text-slate-800">{wo.woNumber}</span>
            <div className="flex items-center gap-1.5 flex-shrink-0">
              {wo.workType && (
                <span className={`text-[9px] font-bold px-1.5 py-0.5 rounded-full ${
                  wo.workType === 'Thermo' ? 'bg-amber-100 text-amber-700' : 'bg-blue-100 text-blue-700'
                }`}>{wo.workType}</span>
              )}
              <StatusBadge status={wo.status} />
            </div>
          </div>
          <p className="text-xs text-slate-500 mt-0.5 truncate">{wo.location || '\u2014'}</p>
          {streets && <p className="text-[11px] text-slate-400 truncate">{streets}</p>}
        </div>

        {/* Expand chevron */}
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="2" strokeLinecap="round"
             className={`flex-shrink-0 mt-1 transition-transform duration-200 ${expanded ? 'rotate-180' : ''}`}>
          <path d="M6 9l6 6 6-6" />
        </svg>
      </button>

      {/* Expanded action bar */}
      {expanded && (
        <div className="px-4 pb-4 sm:pb-3 pt-0 flex flex-wrap gap-2">
          {directionsUrl && (
            <a href={directionsUrl} target="_blank" rel="noopener noreferrer"
               className="flex items-center gap-1.5 px-4 py-2.5 sm:py-2 rounded-xl bg-emerald-600 text-white text-sm font-semibold active:opacity-80 transition-opacity">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round">
                <polygon points="3 11 22 2 13 21 11 13 3 11" />
              </svg>
              Navigate
            </a>
          )}
          <Link to={`/field-report?wo=${encodeURIComponent(wo.woNumber)}`}
                className="flex items-center gap-1.5 px-4 py-2.5 sm:py-2 rounded-xl bg-navy text-white text-sm font-semibold active:opacity-80 transition-opacity">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round">
              <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" />
              <polyline points="14 2 14 8 20 8" />
              <line x1="16" y1="13" x2="8" y2="13" /><line x1="16" y1="17" x2="8" y2="17" />
            </svg>
            Field Report
          </Link>
          {wo.scanFileKey && (
            <button type="button" onClick={onViewPdf}
                    className="flex items-center gap-1.5 px-4 py-2.5 sm:py-2 rounded-xl border border-slate-200 text-slate-700 text-sm font-semibold active:bg-slate-50 transition-colors">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round">
                <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" />
                <polyline points="14 2 14 8 20 8" />
              </svg>
              View Scan
            </button>
          )}
        </div>
      )}
    </div>
  )
}

// ── Map Pin Popover ───────────────────────────────────────────

function MapPopover({ wo, onViewPdf }) {
  const streets = [wo.fromStreet, wo.toStreet].filter(Boolean).join(' \u2192 ')
  const directionsUrl = getDirectionsUrl(wo)

  return (
    <div style={{ minWidth: 220, maxWidth: 280 }} className="space-y-2">
      <div className="flex items-center justify-between gap-2">
        <span className="font-mono font-bold text-navy text-sm">{wo.woNumber}</span>
        <StatusBadge status={wo.status} />
      </div>
      <p className="text-xs text-slate-500">{wo.location || '\u2014'}</p>
      {streets && <p className="text-[11px] text-slate-400">{streets}</p>}

      <div className="flex gap-2 pt-1">
        {directionsUrl && (
          <a href={directionsUrl} target="_blank" rel="noopener noreferrer"
             className="flex-1 text-center text-xs font-bold px-2 py-2.5 rounded-lg bg-emerald-600 text-white active:opacity-80">
            Navigate
          </a>
        )}
        <Link to={`/field-report?wo=${encodeURIComponent(wo.woNumber)}`}
              className="flex-1 text-center text-xs font-bold px-2 py-2.5 rounded-lg bg-navy text-white active:opacity-80">
          Field Report
        </Link>
      </div>
      {wo.scanFileKey && (
        <button type="button" onClick={onViewPdf}
                className="w-full text-center text-xs font-bold px-2 py-2 rounded-lg border border-slate-200 text-slate-600 active:bg-slate-50">
          View Scan
        </button>
      )}
    </div>
  )
}
