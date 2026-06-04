import { useState, useEffect, useMemo, useCallback, useRef } from 'react'
import { Link } from 'react-router-dom'
import { GoogleMap, useJsApiLoader, MarkerF, InfoWindowF } from '@react-google-maps/api'
import StatusBadge from '../components/StatusBadge'
import EditCoordinatesModal from '../components/EditCoordinatesModal'

// Default center = midtown Manhattan-ish. Zoom level fits all 5 boroughs.
const NYC_CENTER = { lat: 40.7128, lng: -74.0060 }
const NYC_ZOOM   = 11

// Width is inline (the lib needs an explicit container), height is a
// Tailwind class so it can shrink on phones — a fixed 600px map dominates
// a small screen. 60vh leaves room to scroll the page past the map.
const MAP_CONTAINER_STYLE = { width: '100%' }
const MAP_HEIGHT_CLASS = 'w-full h-[60vh] sm:h-[600px]'
const MAP_OPTIONS = {
  disableDefaultUI: false,
  streetViewControl: false,
  mapTypeControl: true,
  fullscreenControl: true,
  // Cooperative: one-finger drag scrolls the PAGE, two fingers pan the
  // map, pinch zooms. Without this the map swallows one-finger scroll on
  // mobile and the user can't scroll past it.
  gestureHandling: 'cooperative',
}

// Hex pin colors mirror StatusBadge palette so the legend the user
// already learned on the WO Tracker carries over to the map.
const PIN_COLOR = {
  'received':    '#3b82f6',  // blue-500
  'dispatched':  '#f59e0b',  // amber-500
  'in progress': '#f97316',  // orange-500
}

// Build a Google Maps Symbol icon for a colored teardrop pin. The
// strokeColor flips red on pins with a geocode_warning so admins
// notice them at a glance.
function buildPinIcon(status, hasWarning) {
  const color = PIN_COLOR[(status || '').toLowerCase()] || '#94a3b8'  // slate fallback
  return {
    path: 'M 12,2 C 7.589,2 4,5.589 4,10 c 0,5.5 8,12 8,12 0,0 8,-6.5 8,-12 0,-4.411 -3.589,-8 -8,-8 z',
    fillColor: color,
    fillOpacity: 1,
    strokeColor: hasWarning ? '#dc2626' : '#0f172a',  // red-600 or navy
    strokeWeight: hasWarning ? 2.5 : 1.2,
    scale: 1.6,
    anchor: { x: 12, y: 22 },
  }
}

// Multi-select filter pill bar — same UX as the WO Tracker tab. All
// pill = empty Set; clicking a value toggles it in. Empty Set = no
// filter for that category.
function FilterPills({ label, options, selected, onToggle, onClear }) {
  const allActive = selected.size === 0
  const pillClass = (active) =>
    `text-xs px-3 py-1 rounded-full border font-medium transition-all
     ${active
       ? 'bg-navy text-white border-navy'
       : 'bg-white text-slate-600 border-slate-200 hover:border-navy/40'
     }`
  return (
    <div className="flex items-center gap-2 flex-wrap">
      <span className="text-xs font-semibold text-slate-500">{label}</span>
      <button type="button" onClick={onClear} className={pillClass(allActive)}>All</button>
      {options.map(opt => (
        <button
          key={opt}
          type="button"
          onClick={() => onToggle(opt)}
          className={pillClass(selected.has(opt))}
        >
          {opt}
        </button>
      ))}
    </div>
  )
}

export default function NavTab() {
  const apiKey = import.meta.env.VITE_GOOGLE_MAPS_BROWSER_KEY
  const { isLoaded, loadError } = useJsApiLoader({
    id: 'google-maps-script',
    googleMapsApiKey: apiKey || '',
  })

  const [mapped,    setMapped]    = useState([])
  const [unmapped,  setUnmapped]  = useState([])
  const [loading,   setLoading]   = useState(true)
  const [error,     setError]     = useState('')
  const [activePin, setActivePin] = useState(null)  // wo object whose InfoWindow is open
  const [editingWo, setEditingWo] = useState(null)  // wo passed to EditCoordinatesModal

  const [contFilt,    setContFilt]    = useState(() => new Set())
  const [boroughFilt, setBoroughFilt] = useState(() => new Set())

  // Multi-select toggle / clear helpers — new Set on every change so
  // useMemo re-runs.
  const makeToggle = (setter) => (opt) => setter(prev => {
    const next = new Set(prev)
    if (next.has(opt)) next.delete(opt)
    else              next.add(opt)
    return next
  })

  // Fetch fresh data on tab mount + on every manual refresh.
  const load = useCallback(async () => {
    setLoading(true)
    setError('')
    try {
      const res = await fetch('/api/wos/map')
      const data = await res.json()
      if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`)
      setMapped(Array.isArray(data.mapped) ? data.mapped : [])
      setUnmapped(Array.isArray(data.unmapped) ? data.unmapped : [])
    } catch (e) {
      setError(e.message || 'Failed to load map data')
    } finally {
      setLoading(false)
    }
  }, [])
  useEffect(() => { load() }, [load])

  // ── Lock page-level zoom while on the Nav tab (except on the map) ──
  // In the field it's easy to accidentally pinch-zoom the *page* outside
  // the map; once that happens the 60vh map fills the screen and it's
  // hard to pinch back out. iOS Safari ignores `user-scalable=no` in the
  // viewport meta, so the only reliable fix is to cancel Safari's pinch
  // gesture events (gesturestart/gesturechange) — but ONLY when the
  // gesture lands outside the map, so the map's own two-finger zoom
  // (gestureHandling: 'cooperative') keeps working. Listeners are scoped
  // to this tab's lifetime: switching tabs unmounts NavTab and restores
  // normal page zoom everywhere else.
  const mapCardRef = useRef(null)
  useEffect(() => {
    const inMap = (target) => !!mapCardRef.current && mapCardRef.current.contains(target)
    const blockOutsideMap = (e) => { if (!inMap(e.target)) e.preventDefault() }
    // touchmove with >1 finger = pinch on engines that don't fire the
    // Safari gesture events; guard the map the same way.
    const blockPinchMove = (e) => {
      if (e.touches && e.touches.length > 1 && !inMap(e.target)) e.preventDefault()
    }
    document.addEventListener('gesturestart',  blockOutsideMap, { passive: false })
    document.addEventListener('gesturechange', blockOutsideMap, { passive: false })
    document.addEventListener('touchmove',     blockPinchMove,  { passive: false })
    return () => {
      document.removeEventListener('gesturestart',  blockOutsideMap)
      document.removeEventListener('gesturechange', blockOutsideMap)
      document.removeEventListener('touchmove',     blockPinchMove)
    }
  }, [])

  // Unique filter option lists derived from the union of mapped +
  // unmapped (so admins can filter to see WOs that need geocoding
  // for a specific contractor too).
  const allWOs = useMemo(() => [...mapped, ...unmapped], [mapped, unmapped])
  const contractors = useMemo(
    () => [...new Set(allWOs.map(w => w.contractor).filter(Boolean))].sort(),
    [allWOs]
  )
  const boroughs = useMemo(
    () => [...new Set(allWOs.map(w => w.borough).filter(Boolean))].sort(),
    [allWOs]
  )

  // Filtered views — empty Set = no filter (matches everything).
  const filteredMapped = useMemo(() => mapped.filter(w =>
    (contFilt.size    === 0 || contFilt.has(w.contractor)) &&
    (boroughFilt.size === 0 || boroughFilt.has(w.borough))
  ), [mapped, contFilt, boroughFilt])
  const filteredUnmapped = useMemo(() => unmapped.filter(w =>
    (contFilt.size    === 0 || contFilt.has(w.contractor)) &&
    (boroughFilt.size === 0 || boroughFilt.has(w.borough))
  ), [unmapped, contFilt, boroughFilt])

  // Center / fit the map to the filtered pins once they load. Falls
  // back to NYC center when there's nothing to fit.
  const onMapLoad = useCallback((map) => {
    if (filteredMapped.length === 0) return
    const bounds = new window.google.maps.LatLngBounds()
    filteredMapped.forEach(w => bounds.extend({ lat: w.lat, lng: w.lng }))
    map.fitBounds(bounds, 64)  // 64px padding
  }, [filteredMapped])

  if (!apiKey) {
    return (
      <div className="card p-6 space-y-2">
        <p className="text-sm font-bold text-red-700">Missing VITE_GOOGLE_MAPS_BROWSER_KEY</p>
        <p className="text-xs text-slate-500">
          Set the browser-restricted Google Maps key in Railway env vars to enable the map.
        </p>
      </div>
    )
  }
  if (loadError) {
    return (
      <div className="card p-6 space-y-2">
        <p className="text-sm font-bold text-red-700">Failed to load Google Maps script</p>
        <p className="text-xs text-slate-500">{String(loadError.message || loadError)}</p>
      </div>
    )
  }

  return (
    <div className="space-y-4">
      {/* Top bar — filters + counts + manual refresh */}
      <div className="card p-4 space-y-3">
        <div className="flex items-center justify-between flex-wrap gap-2">
          <div className="space-y-2 flex-1">
            <FilterPills
              label="Contractor"
              options={contractors}
              selected={contFilt}
              onToggle={makeToggle(setContFilt)}
              onClear={() => setContFilt(new Set())}
            />
            <FilterPills
              label="Borough"
              options={boroughs}
              selected={boroughFilt}
              onToggle={makeToggle(setBoroughFilt)}
              onClear={() => setBoroughFilt(new Set())}
            />
          </div>
          <div className="flex items-center gap-2">
            <p className="text-xs text-slate-500">
              <span className="font-bold text-navy">{filteredMapped.length}</span> mapped
              {filteredUnmapped.length > 0 && (
                <> · <span className="font-bold text-amber-700">{filteredUnmapped.length}</span> need coords</>
              )}
            </p>
            <button
              type="button"
              onClick={load}
              disabled={loading}
              className="btn-outline text-xs px-3 py-1.5 disabled:opacity-50"
            >
              <span className={loading ? 'animate-spin inline-block' : ''}>↻</span> Refresh
            </button>
          </div>
        </div>

        {error && (
          <p className="text-xs text-red-700 bg-red-50 border border-red-200 rounded p-2">⚠ {error}</p>
        )}
      </div>

      {/* Map */}
      <div ref={mapCardRef} className="card overflow-hidden">
        {!isLoaded ? (
          <div className={`flex items-center justify-center ${MAP_HEIGHT_CLASS}`}>
            <div className="w-9 h-9 border-[3px] border-slate-200 border-t-navy rounded-full animate-spin" />
          </div>
        ) : (
          <GoogleMap
            mapContainerStyle={MAP_CONTAINER_STYLE}
            mapContainerClassName={MAP_HEIGHT_CLASS}
            center={NYC_CENTER}
            zoom={NYC_ZOOM}
            options={MAP_OPTIONS}
            onLoad={onMapLoad}
          >
            {filteredMapped.map(w => (
              <MarkerF
                key={w.wo_id}
                position={{ lat: w.lat, lng: w.lng }}
                icon={buildPinIcon(w.status, !!w.geocode_warning)}
                onClick={() => setActivePin(w)}
              />
            ))}

            {activePin && (
              <InfoWindowF
                position={{ lat: activePin.lat, lng: activePin.lng }}
                onCloseClick={() => setActivePin(null)}
              >
                <PinPopover
                  wo={activePin}
                  onEditCoords={() => { setEditingWo(activePin); setActivePin(null) }}
                />
              </InfoWindowF>
            )}
          </GoogleMap>
        )}
      </div>

      {/* Needs coordinates panel — every active WO without a usable
          pin. Click "Set Coordinates" to open the manual-entry modal. */}
      {filteredUnmapped.length > 0 && (
        <div className="card p-4 space-y-3">
          <div className="flex items-baseline justify-between">
            <p className="section-label">Needs coordinates</p>
            <span className="text-xs text-slate-400">{filteredUnmapped.length} WO{filteredUnmapped.length === 1 ? '' : 's'}</span>
          </div>
          <p className="text-xs text-slate-500">
            These WOs are active but don't have a pin yet — either the geocoder couldn't
            resolve them or the cluster check flagged the result. Set coordinates manually
            to surface them on the map.
          </p>
          <ul className="divide-y divide-slate-100">
            {filteredUnmapped.map(w => (
              <li key={w.wo_id} className="py-2 flex items-center justify-between gap-3 flex-wrap">
                <div className="flex items-center gap-2 flex-wrap text-sm">
                  <Link
                    to={`/field-report?wo=${encodeURIComponent(w.wo_id)}`}
                    className="font-mono font-bold text-navy hover:underline"
                  >
                    {w.wo_id}
                  </Link>
                  <span className="text-slate-500">{w.contractor || '—'}</span>
                  <span className="bg-slate-100 text-slate-600 rounded px-1.5 py-0.5 text-[10px] font-semibold uppercase">
                    {w.borough}
                  </span>
                  <span className="text-slate-500">{w.location || '—'}</span>
                  {w.geocode_warning && (
                    <span className="text-amber-700 text-xs italic">⚠ {w.geocode_warning}</span>
                  )}
                </div>
                <button
                  type="button"
                  onClick={() => setEditingWo(w)}
                  className="btn-outline text-xs px-3 py-1 w-full sm:w-auto"
                >
                  Set Coordinates…
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}

      {editingWo && (
        <EditCoordinatesModal
          wo={editingWo}
          onClose={() => setEditingWo(null)}
          onSaved={() => {
            setEditingWo(null)
            load()
          }}
        />
      )}
    </div>
  )
}

// Pin's InfoWindow body. Renders every field the user listed
// (WO #, Contractor, Contract #, Borough, Location, From → To,
// Priority, Due Date, Status, marking-item count) plus a link to
// Field Report and an Edit Coordinates trigger.
function PinPopover({ wo, onEditCoords }) {
  const dueDate = wo.due_date || '—'
  const fromTo = [wo.from_street, wo.to_street].filter(Boolean).join(' → ') || '—'
  const total  = wo.marking_item_count || 0
  const done   = wo.marking_completed_count || 0

  return (
    <div className="space-y-2 max-w-[280px]" style={{ minWidth: 220 }}>
      <div className="flex items-center justify-between gap-2">
        <Link
          to={`/field-report?wo=${encodeURIComponent(wo.wo_id)}`}
          className="font-mono font-bold text-navy hover:underline text-sm"
        >
          {wo.wo_id}
        </Link>
        <StatusBadge status={wo.status} />
      </div>

      {wo.geocode_warning && (
        <p className="text-[11px] text-amber-700 bg-amber-50 border border-amber-200 rounded px-2 py-1 leading-snug">
          ⚠ {wo.geocode_warning}
        </p>
      )}

      <table className="text-xs w-full">
        <tbody>
          <Row label="Contractor"   value={wo.contractor} />
          <Row label="Contract #"   value={wo.contract_num} />
          <Row label="Borough"      value={wo.borough} />
          <Row label="Location"     value={wo.location} />
          <Row label="From → To"    value={fromTo} />
          {wo.priority && <Row label="Priority" value={wo.priority} />}
          <Row label="Due Date"     value={dueDate} />
          <Row label="Markings"     value={total > 0 ? `${done} / ${total} completed` : '—'} />
        </tbody>
      </table>

      <div className="flex gap-2 pt-1">
        <Link
          to={`/field-report?wo=${encodeURIComponent(wo.wo_id)}`}
          className="flex-1 text-center text-xs font-bold px-2 py-1.5 rounded-lg
                     bg-navy text-white hover:opacity-90"
        >
          View Field Report
        </Link>
        <button
          type="button"
          onClick={onEditCoords}
          className="text-xs font-bold px-2 py-1.5 rounded-lg
                     bg-white text-slate-700 border border-slate-200 hover:bg-slate-50"
        >
          Edit Coords
        </button>
      </div>
    </div>
  )
}

function Row({ label, value }) {
  return (
    <tr>
      <td className="text-slate-400 font-semibold pr-2 py-0.5 align-top whitespace-nowrap">{label}</td>
      <td className="text-slate-700 py-0.5">{value || '—'}</td>
    </tr>
  )
}
