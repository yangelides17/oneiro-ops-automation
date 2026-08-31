import { useState, useEffect, useMemo, useCallback } from 'react'
import { GoogleMap, useJsApiLoader, MarkerF, InfoWindowF } from '@react-google-maps/api'
import StatusBadge from '../components/StatusBadge'

// ─── Map Config ───────────────────────────────────────────────

const NYC_CENTER = { lat: 40.7128, lng: -74.0060 }
const NYC_ZOOM = 11
const MAP_OPTIONS = {
  disableDefaultUI: true,
  zoomControl: true,
  mapTypeControl: false,
  streetViewControl: false,
  fullscreenControl: false,
}

// Crew colors — auto-assigned to crew leaders in order
const CREW_COLORS = [
  '#3b82f6', // blue
  '#22c55e', // green
  '#f59e0b', // amber
  '#ef4444', // red
  '#8b5cf6', // purple
  '#06b6d4', // cyan
  '#f97316', // orange
  '#ec4899', // pink
]
const UNASSIGNED_COLOR = '#94a3b8' // slate-400

// ─── Pin rendering (matches NavTab exactly) ───────────────────

const PIN_SHAPE = {
  teardrop: {
    path: 'M 12,2 C 7.589,2 4,5.589 4,10 c 0,5.5 8,12 8,12 0,0 8,-6.5 8,-12 0,-4.411 -3.589,-8 -8,-8 z',
    anchor: { x: 12, y: 22 },
  },
  square: {
    path: 'M 7,2 L 17,2 Q 20,2 20,5 L 20,13 Q 20,16 17,16 L 15,16 L 12,22 L 9,16 L 7,16 Q 4,16 4,13 L 4,5 Q 4,2 7,2 Z',
    anchor: { x: 12, y: 22 },
  },
}
const PIN_SCALE = 1.6

function resolveColor(wo, crewColorMap) {
  const s = String(wo.status || '').toLowerCase()
  if (s === 'completed') return '#d1d5db'
  if (wo.assignedTo && crewColorMap.has(wo.assignedTo)) return crewColorMap.get(wo.assignedTo)
  if (wo.assignedTo) return CREW_COLORS[0]
  return UNASSIGNED_COLOR
}

function buildAssignPin(wo, crewColorMap) {
  const color = resolveColor(wo, crewColorMap)
  const isPT = String(wo.woId || wo.woNumber || '').toUpperCase().startsWith('PT')
  const shape = isPT ? PIN_SHAPE.square : PIN_SHAPE.teardrop

  return {
    path: shape.path,
    fillColor: color,
    fillOpacity: 1,
    strokeColor: wo.assignedTo ? '#0f172a' : '#64748b',
    strokeWeight: wo.assignedTo ? 1.5 : 0.8,
    scale: PIN_SCALE,
    anchor: shape.anchor,
  }
}

// ─── Component ────────────────────────────────────────────────

export default function WOAssignment() {
  const apiKey = import.meta.env.VITE_GOOGLE_MAPS_BROWSER_KEY
  const { isLoaded } = useJsApiLoader({ id: 'google-maps-script', googleMapsApiKey: apiKey || '' })

  const [wos, setWos] = useState([])
  const [crewLeaders, setCrewLeaders] = useState([])
  const [loading, setLoading] = useState(true)
  const [activePin, setActivePin] = useState(null)
  const [workTypeFilter, setWorkTypeFilter] = useState('')
  const [mobileView, setMobileView] = useState('panel') // 'panel' | 'map'

  const loadData = useCallback(async () => {
    setLoading(true)
    try {
      const [mapRes, usersRes] = await Promise.all([
        fetch('/api/wos/map').then(r => r.json()),
        fetch('/api/settings/users').then(r => r.json()),
      ])
      const allWos = [...(mapRes.mapped || []), ...(mapRes.unmapped || [])]
      setWos(allWos)
      const leaders = (Array.isArray(usersRes) ? usersRes : [])
        .filter(u => u.role === 'crew' || u.role === 'foreman')
      setCrewLeaders(leaders)
    } catch (e) {
      console.error('Failed to load assignment data:', e)
    } finally {
      setLoading(false)
    }
  }, [])

  useEffect(() => { loadData() }, [loadData])

  // Build crew color map
  const crewColorMap = useMemo(() => {
    const map = new Map()
    crewLeaders.forEach((cl, i) => {
      map.set(cl.id, CREW_COLORS[i % CREW_COLORS.length])
    })
    wos.forEach(wo => {
      if (wo.assignedTo && !map.has(wo.assignedTo)) {
        map.set(wo.assignedTo, CREW_COLORS[map.size % CREW_COLORS.length])
      }
    })
    return map
  }, [crewLeaders, wos])

  // Filter WOs by work type
  const filteredWos = useMemo(() => {
    if (!workTypeFilter) return wos
    return wos.filter(w => (w.workType || '').toLowerCase() === workTypeFilter.toLowerCase())
  }, [wos, workTypeFilter])

  const mappedWos = useMemo(() => filteredWos.filter(w => w.lat && w.lng), [filteredWos])

  // Group WOs by assignment
  const grouped = useMemo(() => {
    const unassigned = []
    const byCrewMap = new Map()
    crewLeaders.forEach(cl => byCrewMap.set(cl.id, { leader: cl, wos: [] }))

    for (const wo of filteredWos) {
      if (!wo.assignedTo) {
        unassigned.push(wo)
      } else if (byCrewMap.has(wo.assignedTo)) {
        byCrewMap.get(wo.assignedTo).wos.push(wo)
      } else {
        unassigned.push(wo)
      }
    }

    return {
      unassigned,
      crews: Array.from(byCrewMap.values()).filter(c => c.wos.length > 0),
    }
  }, [filteredWos, crewLeaders])

  const assignWo = async (woId, userId) => {
    await fetch('/api/wos/assign', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ woIds: [woId], userId }),
    })
    setActivePin(null)
    loadData()
  }

  const unassignWo = (woId) => assignWo(woId, null)

  return (
    <div className="h-[calc(100vh-56px)] flex flex-col">
      {/* ── Top bar ────────────────────────────────────────────── */}
      <div className="flex items-center justify-between gap-2 px-4 py-2 border-b border-slate-200 bg-white flex-shrink-0">
        <h1 className="text-lg font-black text-navy">Assign WOs</h1>
        <div className="flex items-center gap-2">
          {/* Mobile view toggle */}
          <div className="sm:hidden flex bg-slate-100 rounded-lg p-0.5">
            <button
              onClick={() => setMobileView('panel')}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-all ${
                mobileView === 'panel' ? 'bg-white text-navy shadow-sm' : 'text-slate-500'
              }`}
            >
              Panel
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
          <select value={workTypeFilter} onChange={e => setWorkTypeFilter(e.target.value)}
                  className="field-input text-sm w-auto">
            <option value="">All Types</option>
            <option value="Thermo">Thermo</option>
            <option value="MMA">MMA</option>
          </select>
          <button onClick={loadData} disabled={loading}
                  className="p-2 rounded-lg hover:bg-slate-100 active:bg-slate-200 transition-colors"
                  title="Refresh">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"
                 className={`text-slate-500 ${loading ? 'animate-spin' : ''}`}>
              <path d="M21 12a9 9 0 11-6.219-8.56" /><path d="M21 3v6h-6" />
            </svg>
          </button>
        </div>
      </div>

      {/* ── Main split: map + panel ────────────────────────────── */}
      <div className="flex-1 flex flex-col sm:flex-row min-h-0">
        {/* Map — full screen on mobile (map view), flex-1 on desktop */}
        <div className={`
          ${mobileView === 'map' ? 'flex-1' : 'hidden'}
          sm:block sm:flex-1 min-h-0 min-w-0
        `}>
          {isLoaded ? (
            <GoogleMap
              mapContainerStyle={{ width: '100%', height: '100%' }}
              center={NYC_CENTER}
              zoom={NYC_ZOOM}
              options={MAP_OPTIONS}
            >
              {mappedWos.map(w => (
                <MarkerF
                  key={w.woId || w.id}
                  position={{ lat: w.lat, lng: w.lng }}
                  icon={buildAssignPin(w, crewColorMap)}
                  onClick={() => setActivePin(w)}
                />
              ))}
              {activePin && (
                <InfoWindowF
                  position={{ lat: activePin.lat, lng: activePin.lng }}
                  onCloseClick={() => setActivePin(null)}
                >
                  <AssignPopover
                    wo={activePin}
                    crewLeaders={crewLeaders}
                    crewColorMap={crewColorMap}
                    onAssign={assignWo}
                    onUnassign={unassignWo}
                  />
                </InfoWindowF>
              )}
            </GoogleMap>
          ) : (
            <div className="flex items-center justify-center h-full">
              <div className="w-9 h-9 border-[3px] border-slate-200 border-t-navy rounded-full animate-spin" />
            </div>
          )}
        </div>

        {/* Assignment panel — full screen on mobile (panel view), fixed width on desktop */}
        <div className={`
          ${mobileView === 'panel' ? 'flex-1' : 'hidden'}
          sm:block sm:w-[340px] sm:flex-none
          border-l border-slate-200 bg-white overflow-y-auto min-h-0
        `}>
          {/* Legend */}
          <div className="px-3 py-2 border-b border-slate-100">
            <div className="flex flex-wrap gap-2 text-[10px]">
              <span className="flex items-center gap-1">
                <span className="w-2.5 h-2.5 rounded-full" style={{ background: UNASSIGNED_COLOR }} />
                Unassigned
              </span>
              {grouped.crews.map(c => (
                <span key={c.leader.id} className="flex items-center gap-1">
                  <span className="w-2.5 h-2.5 rounded-full" style={{ background: crewColorMap.get(c.leader.id) }} />
                  {c.leader.name}
                </span>
              ))}
            </div>
          </div>

          {/* Unassigned section */}
          <div className="border-b border-slate-100">
            <div className="px-3 py-2 bg-slate-50 flex items-center justify-between">
              <span className="text-xs font-bold text-slate-500">
                Unassigned ({grouped.unassigned.length})
              </span>
            </div>
            {grouped.unassigned.length === 0 && (
              <p className="px-3 py-3 text-xs text-slate-400 italic">All work orders assigned</p>
            )}
            {grouped.unassigned.map(wo => (
              <WORow key={wo.id} wo={wo}
                     crewLeaders={crewLeaders} crewColorMap={crewColorMap}
                     onAssign={assignWo}
                     onClickPin={() => { setActivePin(wo); setMobileView('map') }} />
            ))}
          </div>

          {/* Crew sections */}
          {grouped.crews.map(c => (
            <div key={c.leader.id} className="border-b border-slate-100">
              <div className="px-3 py-2 flex items-center justify-between"
                   style={{ backgroundColor: crewColorMap.get(c.leader.id) + '15' }}>
                <span className="flex items-center gap-2">
                  <span className="w-3 h-3 rounded-full" style={{ background: crewColorMap.get(c.leader.id) }} />
                  <span className="text-xs font-bold text-slate-700">{c.leader.name}</span>
                </span>
                <span className="text-[10px] text-slate-400">{c.wos.length} WO{c.wos.length === 1 ? '' : 's'}</span>
              </div>
              {c.wos.map(wo => (
                <WORow key={wo.id} wo={wo}
                       onUnassign={() => unassignWo(wo.id)}
                       onClickPin={() => { setActivePin(wo); setMobileView('map') }} />
              ))}
            </div>
          ))}

          {crewLeaders.length === 0 && (
            <div className="px-3 py-4 text-xs text-slate-400 italic text-center">
              No crew leaders found. Invite users with the "crew" or "foreman" role in Settings.
            </div>
          )}
        </div>
      </div>
    </div>
  )
}

// ─── WO Row in assignment panel ───────────────────────────────

function WORow({ wo, onUnassign, onClickPin, crewLeaders, crewColorMap, onAssign }) {
  const [showAssign, setShowAssign] = useState(false)
  const loc = wo.location || ''
  const streets = [wo.fromStreet, wo.toStreet].filter(Boolean).join(' \u2192 ')

  return (
    <div className="border-b border-slate-50">
      <div
        className="flex items-center gap-2 px-3 py-2.5 sm:py-2 hover:bg-slate-50 active:bg-slate-100 cursor-pointer transition-colors"
        onClick={onClickPin}
      >
        <div className="flex-1 min-w-0">
          <p className="text-xs font-bold text-slate-700 truncate">{wo.woId || wo.woNumber}</p>
          <p className="text-[10px] text-slate-400 truncate">{loc}{streets ? ` (${streets})` : ''}</p>
        </div>
        {wo.workType && (
          <span className={`text-[9px] font-bold px-1.5 py-0.5 rounded-full flex-shrink-0 ${
            wo.workType === 'Thermo' ? 'bg-amber-100 text-amber-700' : 'bg-blue-100 text-blue-700'
          }`}>
            {wo.workType}
          </span>
        )}
        {/* Assign button (for unassigned WOs on mobile) */}
        {crewLeaders && crewLeaders.length > 0 && (
          <button
            type="button"
            onClick={e => { e.stopPropagation(); setShowAssign(v => !v) }}
            className="text-[10px] font-bold px-2 py-1 rounded-lg bg-navy/10 text-navy flex-shrink-0 active:bg-navy/20"
          >
            Assign
          </button>
        )}
        {onUnassign && (
          <button
            type="button"
            onClick={e => { e.stopPropagation(); onUnassign() }}
            className="w-7 h-7 flex items-center justify-center rounded-lg text-slate-300 hover:text-red-500 hover:bg-red-50 active:bg-red-100 text-sm flex-shrink-0 transition-colors"
            title="Unassign"
          >
            &times;
          </button>
        )}
      </div>
      {/* Inline assign picker */}
      {showAssign && crewLeaders && (
        <div className="px-3 pb-2 flex flex-wrap gap-1.5">
          {crewLeaders.map(cl => (
            <button
              key={cl.id}
              type="button"
              onClick={() => { onAssign(wo.id, cl.id); setShowAssign(false) }}
              className="text-[11px] font-bold px-3 py-1.5 rounded-lg border active:opacity-80"
              style={{
                borderColor: crewColorMap.get(cl.id),
                color: crewColorMap.get(cl.id),
                backgroundColor: crewColorMap.get(cl.id) + '15',
              }}
            >
              {cl.name}
            </button>
          ))}
        </div>
      )}
    </div>
  )
}

// ─── Pin Popover with assign dropdown ─────────────────────────

function AssignPopover({ wo, crewLeaders, crewColorMap, onAssign, onUnassign }) {
  const fromTo = [wo.fromStreet, wo.toStreet].filter(Boolean).join(' \u2192 ') || '\u2014'

  return (
    <div className="space-y-2" style={{ minWidth: 220, maxWidth: 280 }}>
      <div className="flex items-center justify-between gap-2">
        <span className="font-mono font-bold text-navy text-sm">{wo.woId || wo.woNumber}</span>
        <StatusBadge status={wo.status} />
      </div>

      <div className="text-xs space-y-0.5">
        <p><span className="text-slate-400">Location:</span> {wo.location || '\u2014'}</p>
        <p><span className="text-slate-400">From \u2192 To:</span> {fromTo}</p>
        {wo.workType && <p><span className="text-slate-400">Type:</span> {wo.workType}</p>}
      </div>

      {wo.assignedTo ? (
        <div className="flex items-center justify-between pt-1">
          <span className="text-xs">
            <span className="text-slate-400">Assigned to:</span>{' '}
            <span className="font-bold" style={{ color: crewColorMap.get(wo.assignedTo) }}>
              {wo.assignedToName || 'Unknown'}
            </span>
          </span>
          <button
            type="button"
            onClick={() => onUnassign(wo.id)}
            className="text-[10px] font-bold px-2 py-1.5 rounded-lg text-red-600 bg-red-50 hover:bg-red-100 active:bg-red-200"
          >
            Unassign
          </button>
        </div>
      ) : (
        <div className="pt-1">
          <p className="text-[10px] text-slate-400 mb-1.5">Assign to:</p>
          <div className="flex flex-wrap gap-1.5">
            {crewLeaders.map(cl => (
              <button
                key={cl.id}
                type="button"
                onClick={() => onAssign(wo.id, cl.id)}
                className="text-[11px] font-bold px-3 py-1.5 rounded-lg border active:opacity-80"
                style={{
                  borderColor: crewColorMap.get(cl.id),
                  color: crewColorMap.get(cl.id),
                  backgroundColor: crewColorMap.get(cl.id) + '15',
                }}
              >
                {cl.name}
              </button>
            ))}
          </div>
        </div>
      )}
    </div>
  )
}
