import { useEffect, useRef, useState } from 'react'

/**
 * EditCoordinatesModal — manual lat/lng entry for the Nav tab.
 *
 * Two input modes share the same form:
 *   - Paste a Google Maps URL → we regex-parse the coords
 *   - Type lat / lng directly
 *
 * Both update the same hidden state. Save POSTs to
 * /api/wo/:id/coordinates. Parent receives the saved coords via
 * onSaved so it can refresh the map without re-fetching.
 *
 * Props:
 *   wo       — { wo_id, location, lat, lng, geocode_warning } (existing coords pre-fill if set)
 *   onSaved  — (newCoords: { lat, lng }) => void
 *   onClose  — () => void
 */
export default function EditCoordinatesModal({ wo, onSaved, onClose }) {
  const [lat, setLat]               = useState(wo?.lat != null ? String(wo.lat) : '')
  const [lng, setLng]               = useState(wo?.lng != null ? String(wo.lng) : '')
  const [url, setUrl]               = useState('')
  const [submitting, setSubmitting] = useState(false)
  const [error, setError]           = useState('')
  const firstInputRef               = useRef(null)

  useEffect(() => {
    const handler = (e) => { if (e.key === 'Escape' && !submitting) onClose?.() }
    window.addEventListener('keydown', handler)
    setTimeout(() => firstInputRef.current?.focus(), 0)
    return () => window.removeEventListener('keydown', handler)
  }, [onClose, submitting])

  if (!wo) return null

  // Parse coords out of a pasted Google Maps URL. Handles the two
  // most common formats:
  //   https://www.google.com/maps/.../@40.7128,-74.0060,17z
  //   https://www.google.com/maps?q=40.7128,-74.0060
  //   https://maps.app.goo.gl/...     (these don't contain coords —
  //     user has to "Share" → "Copy link" and let it redirect first,
  //     OR right-click → "What's here?" and copy from the popup)
  // Falls back to "first two comma-separated decimals" in the string.
  const parseFromUrl = (raw) => {
    const s = String(raw || '').trim()
    if (!s) return null
    // @lat,lng pattern
    let m = s.match(/@(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/)
    if (m) return { lat: m[1], lng: m[2] }
    // ?q=lat,lng or &q=lat,lng
    m = s.match(/[?&]q=(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/)
    if (m) return { lat: m[1], lng: m[2] }
    // Plain "lat, lng" (e.g. copied from the "What's here?" popup)
    m = s.match(/(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)/)
    if (m) return { lat: m[1], lng: m[2] }
    return null
  }

  const handleUrlBlur = () => {
    if (!url) return
    const parsed = parseFromUrl(url)
    if (parsed) {
      setLat(parsed.lat)
      setLng(parsed.lng)
      setError('')
    } else {
      setError('Could not parse coords from that URL — try the lat/lng fields directly.')
    }
  }

  const save = async () => {
    setError('')
    const la = parseFloat(lat)
    const ln = parseFloat(lng)
    if (isNaN(la) || la < -90  || la > 90)  { setError('Latitude must be a number between -90 and 90.');   return }
    if (isNaN(ln) || ln < -180 || ln > 180) { setError('Longitude must be a number between -180 and 180.'); return }
    setSubmitting(true)
    try {
      const res = await fetch(`/api/wo/${encodeURIComponent(wo.wo_id)}/coordinates`, {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ lat: la, lng: ln }),
      })
      const body = await res.json().catch(() => ({}))
      if (!res.ok || body.error) throw new Error(body.error || `HTTP ${res.status}`)
      onSaved?.({ lat: la, lng: ln })
      onClose?.()
    } catch (e) {
      setError(e.message || 'Failed to save coordinates')
    } finally {
      setSubmitting(false)
    }
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4"
      style={{ backgroundColor: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(2px)' }}
      onClick={() => !submitting && onClose?.()}
    >
      <div
        className="bg-white rounded-2xl shadow-2xl w-full max-w-md p-6 space-y-4"
        onClick={e => e.stopPropagation()}
      >
        <div className="text-center space-y-1.5">
          <h2 className="text-lg font-black text-navy">Set Coordinates</h2>
          <p className="text-slate-500 text-sm">
            <span className="font-mono font-bold text-navy">{wo.wo_id}</span>
            {wo.location && <> — {wo.location}</>}
          </p>
        </div>

        {wo.geocode_warning && (
          <div className="bg-amber-50 border border-amber-200 rounded-lg p-3 text-xs text-amber-800">
            <strong>Existing warning:</strong> {wo.geocode_warning}
          </div>
        )}

        <div className="space-y-2">
          <label className="block">
            <span className="text-[11px] font-bold uppercase tracking-wider text-slate-500">
              Paste a Google Maps URL or "What's here?" coords
            </span>
            <input
              ref={firstInputRef}
              type="text"
              value={url}
              onChange={e => setUrl(e.target.value)}
              onBlur={handleUrlBlur}
              placeholder="https://www.google.com/maps/@40.7128,-74.0060,17z"
              className="field-input text-sm mt-1"
              disabled={submitting}
            />
            <p className="text-[10px] text-slate-400 mt-1">
              Tip: right-click a spot in Google Maps → "What's here?" → click the lat/lng
              popup → coords copy to clipboard.
            </p>
          </label>

          <div className="grid grid-cols-2 gap-2 pt-1">
            <label className="block">
              <span className="text-[11px] font-bold uppercase tracking-wider text-slate-500">Latitude</span>
              <input
                type="text"
                value={lat}
                onChange={e => setLat(e.target.value)}
                placeholder="40.7128"
                className="field-input text-sm mt-1"
                disabled={submitting}
              />
            </label>
            <label className="block">
              <span className="text-[11px] font-bold uppercase tracking-wider text-slate-500">Longitude</span>
              <input
                type="text"
                value={lng}
                onChange={e => setLng(e.target.value)}
                placeholder="-74.0060"
                className="field-input text-sm mt-1"
                disabled={submitting}
              />
            </label>
          </div>
        </div>

        {error && (
          <p className="text-xs text-red-600 font-semibold text-center">{error}</p>
        )}

        <div className="flex flex-col gap-2 pt-1">
          <button
            type="button"
            disabled={submitting || !lat || !lng}
            onClick={save}
            className="w-full py-3 rounded-xl font-bold text-sm bg-navy text-white
                       hover:opacity-90 active:opacity-80 transition-all
                       disabled:opacity-40 disabled:cursor-not-allowed"
          >
            {submitting ? 'Saving…' : 'Save Coordinates'}
          </button>
          <button
            type="button"
            disabled={submitting}
            onClick={onClose}
            className="w-full py-3 rounded-xl font-bold text-sm bg-slate-100
                       text-slate-600 hover:bg-slate-200 transition-all
                       disabled:opacity-40"
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  )
}
