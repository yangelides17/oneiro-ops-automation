import { useState, useEffect, useMemo, useCallback } from 'react'
import { Link } from 'react-router-dom'
import { GoogleMap, useJsApiLoader, MarkerF, InfoWindowF } from '@react-google-maps/api'
import StatusBadge from '../components/StatusBadge'
import PdfViewer from '../components/PdfViewer'

const NYC_CENTER = { lat: 40.7128, lng: -74.0060 }
const NYC_ZOOM = 11
const MAP_OPTIONS = {
  disableDefaultUI: true,
  zoomControl: true,
  mapTypeControl: false,
  streetViewControl: false,
  fullscreenControl: false,
}

const PIN_COLORS = {
  received:    '#3b82f6',
  dispatched:  '#f59e0b',
  in_progress: '#f97316',
  completed:   '#22c55e',
}

function buildPin(status) {
  const s = String(status || '').toLowerCase().replace(/\s+/g, '_')
  const color = PIN_COLORS[s] || '#3b82f6'
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="24" height="36" viewBox="0 0 24 36">
    <path d="M12 0C5.4 0 0 5.4 0 12c0 9 12 24 12 24s12-15 12-24C24 5.4 18.6 0 12 0z" fill="${color}" stroke="#fff" stroke-width="1.5"/>
    <circle cx="12" cy="12" r="5" fill="#fff" opacity="0.9"/>
  </svg>`
  return {
    url: 'data:image/svg+xml;charset=UTF-8,' + encodeURIComponent(svg),
    scaledSize: { width: 24, height: 36 },
    anchor: { x: 12, y: 36 },
  }
}

export default function MyWork() {
  const apiKey = import.meta.env.VITE_GOOGLE_MAPS_BROWSER_KEY
  const { isLoaded } = useJsApiLoader({ id: 'google-maps-script', googleMapsApiKey: apiKey || '' })

  const [wos, setWos] = useState([])
  const [loading, setLoading] = useState(true)
  const [selectedWo, setSelectedWo] = useState(null)
  const [pdfUrl, setPdfUrl] = useState(null)
  const [activePin, setActivePin] = useState(null)

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

  const mappedWos = useMemo(() => wos.filter(w => w.latitude && w.longitude), [wos])

  // Load PDF when a WO is selected
  const selectWo = useCallback((wo) => {
    setSelectedWo(wo)
    setActivePin(wo)
    if (wo.scanFileKey) {
      fetch(`/api/wos/${encodeURIComponent(wo.id)}/files`)
        .then(r => r.json())
        .then(d => setPdfUrl(d.scan?.url || null))
        .catch(() => setPdfUrl(null))
    } else {
      setPdfUrl(null)
    }
  }, [])

  const closePdf = () => { setSelectedWo(null); setPdfUrl(null) }
  const showPdf = !!(selectedWo && pdfUrl)

  // Count by status
  const statusCounts = useMemo(() => {
    const counts = { total: wos.length, pending: 0, inProgress: 0, completed: 0 }
    for (const w of wos) {
      const s = String(w.status || '').toLowerCase().replace(/\s+/g, '_')
      if (s === 'completed') counts.completed++
      else if (s === 'in_progress') counts.inProgress++
      else counts.pending++
    }
    return counts
  }, [wos])

  return (
    <div className="h-[calc(100vh-56px)] flex flex-col">
      {/* Header */}
      <div className="flex items-center justify-between gap-3 px-4 py-2 border-b border-slate-200 bg-white flex-shrink-0">
        <div>
          <h1 className="text-lg font-black text-navy">My Work</h1>
          <p className="text-[11px] text-slate-400">
            {statusCounts.total} work order{statusCounts.total === 1 ? '' : 's'}
            {statusCounts.completed > 0 && <> · <span className="text-green-600">{statusCounts.completed} done</span></>}
            {statusCounts.inProgress > 0 && <> · <span className="text-orange-600">{statusCounts.inProgress} in progress</span></>}
          </p>
        </div>
        <button onClick={loadWork} disabled={loading} className="btn-outline text-xs px-3 py-1.5">
          ↻ Refresh
        </button>
      </div>

      {/* Main content */}
      <div className="flex-1 flex flex-col sm:flex-row min-h-0">
        {/* Map */}
        <div className={`flex-shrink-0 ${showPdf ? 'h-[35vh] sm:h-auto sm:w-[40%]' : 'h-[40vh] sm:h-auto sm:w-[50%]'}`}>
          {isLoaded ? (
            <GoogleMap
              mapContainerStyle={{ width: '100%', height: '100%' }}
              center={mappedWos.length > 0 ? { lat: Number(mappedWos[0].latitude), lng: Number(mappedWos[0].longitude) } : NYC_CENTER}
              zoom={mappedWos.length > 0 ? 13 : NYC_ZOOM}
              options={MAP_OPTIONS}
            >
              {mappedWos.map(w => (
                <MarkerF
                  key={w.id}
                  position={{ lat: Number(w.latitude), lng: Number(w.longitude) }}
                  icon={buildPin(w.status)}
                  onClick={() => { setActivePin(w); selectWo(w) }}
                />
              ))}
              {activePin && activePin.latitude && (
                <InfoWindowF
                  position={{ lat: Number(activePin.latitude), lng: Number(activePin.longitude) }}
                  onCloseClick={() => setActivePin(null)}
                >
                  <div className="space-y-1" style={{ minWidth: 180 }}>
                    <div className="flex items-center justify-between gap-2">
                      <span className="font-mono font-bold text-navy text-sm">{activePin.woNumber}</span>
                      <StatusBadge status={activePin.status} />
                    </div>
                    <p className="text-xs text-slate-500">{activePin.location}</p>
                    <Link
                      to={`/field-report?wo=${encodeURIComponent(activePin.woNumber)}`}
                      className="block text-center text-xs font-bold px-2 py-1.5 rounded-lg bg-navy text-white hover:opacity-90 mt-1"
                    >
                      Open Field Report
                    </Link>
                  </div>
                </InfoWindowF>
              )}
            </GoogleMap>
          ) : (
            <div className="flex items-center justify-center h-full bg-slate-50">
              <div className="w-8 h-8 border-[3px] border-slate-200 border-t-navy rounded-full animate-spin" />
            </div>
          )}
        </div>

        {/* WO list */}
        <div className={`border-l border-slate-200 overflow-y-auto bg-white ${showPdf ? 'hidden sm:block sm:w-[25%]' : 'sm:w-[50%]'}`}>
          {loading && (
            <div className="flex items-center justify-center py-8">
              <div className="w-6 h-6 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
            </div>
          )}
          {!loading && wos.length === 0 && (
            <div className="flex flex-col items-center justify-center py-12 px-4 text-center">
              <span className="text-3xl mb-2">📋</span>
              <p className="text-sm font-semibold text-slate-500">No work orders assigned to you</p>
              <p className="text-xs text-slate-400 mt-1">Your field manager will assign work orders from the WO Assignment tab.</p>
            </div>
          )}
          {wos.map(wo => {
            const isSelected = selectedWo?.id === wo.id
            const streets = [wo.fromStreet, wo.toStreet].filter(Boolean).join(' → ')
            return (
              <button
                key={wo.id}
                type="button"
                onClick={() => selectWo(wo)}
                className={`w-full text-left px-3 py-3 border-b border-slate-100 hover:bg-slate-50 transition-all
                  ${isSelected ? 'bg-navy/5 border-l-2 border-l-navy' : ''}`}
              >
                <div className="flex items-center justify-between gap-2">
                  <span className="text-sm font-bold text-slate-800">{wo.woNumber}</span>
                  <StatusBadge status={wo.status} />
                </div>
                <p className="text-xs text-slate-500 mt-0.5">{wo.location || '—'}</p>
                {streets && <p className="text-[10px] text-slate-400">{streets}</p>}
                <div className="flex items-center gap-2 mt-1">
                  {wo.workType && (
                    <span className={`text-[9px] font-bold px-1.5 py-0.5 rounded-full ${
                      wo.workType === 'Thermo' ? 'bg-amber-100 text-amber-700' : 'bg-blue-100 text-blue-700'
                    }`}>{wo.workType}</span>
                  )}
                  <Link
                    to={`/field-report?wo=${encodeURIComponent(wo.woNumber)}`}
                    onClick={e => e.stopPropagation()}
                    className="text-[10px] font-bold text-navy hover:underline"
                  >
                    Field Report →
                  </Link>
                </div>
              </button>
            )
          })}
        </div>

        {/* PDF viewer */}
        {showPdf && (
          <div className="flex-1 min-w-0 border-l border-slate-200">
            <PdfViewer
              url={pdfUrl}
              filename={`${selectedWo.woNumber} Scan`}
              collapsed={false}
              onToggle={() => {}}
              onClose={closePdf}
              fillParent
            />
          </div>
        )}
      </div>
    </div>
  )
}
