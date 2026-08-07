import { useEffect, useState } from 'react'

/**
 * Polls /api/qb/status and exposes a banner that ONLY renders when
 * the QB connection is unavailable. The banner is mounted above the
 * WO table on the WO Tracker tab; per-row Generate buttons read the
 * same status (via the useQbStatus hook below) to fail closed.
 *
 * Status values from /api/qb/status:
 *   { connected: true,  sandbox: bool }
 *   { connected: false, reason: 'env_missing'   | 'not_authorized' | 'error',
 *     missing?: string[], error?: string, sandbox: bool }
 */

export function useQbStatus() {
  const [status, setStatus] = useState({ loading: true })

  useEffect(() => {
    let cancelled = false
    async function load() {
      try {
        const res = await fetch('/api/qb/status')
        const data = await res.json()
        if (!cancelled) setStatus({ loading: false, ...data })
      } catch (err) {
        if (!cancelled) setStatus({ loading: false, connected: false, reason: 'error', error: err.message })
      }
    }
    load()
    // Poll every 15 minutes, and only while the tab is actually visible.
    //
    // This was every 2 minutes, ungated — 30 executions/hour from every
    // open Dashboard tab, including ones left in the background all day.
    // In the overnight failure dataset, get_qb_refresh_token accounted for
    // 25 of 49 failures: not because it is fragile (it reads two Script
    // Properties) but purely from exposure. Every Apps Script call is an
    // independent roll against a service that intermittently hangs, and
    // this one was rolling more than anything else in the app.
    //
    // What it buys is small: the badge renders NOTHING while QB is
    // connected (see below), so this only needs to notice a disconnect
    // eventually, not promptly. A visibility check also means a truck
    // dashboard left open in a background tab stops polling entirely.
    const tick = () => { if (!document.hidden) load() }
    const id = setInterval(tick, 15 * 60_000)
    const onVis = () => { if (!document.hidden) load() }
    document.addEventListener('visibilitychange', onVis)
    return () => {
      cancelled = true
      clearInterval(id)
      document.removeEventListener('visibilitychange', onVis)
    }
  }, [])

  return status
}

export default function QBStatusBadge({ status }) {
  if (!status || status.loading || status.connected) return null

  const reason = status.reason || 'error'
  const explain =
    reason === 'env_missing'    ? 'QuickBooks integration is not configured on the server.' :
    reason === 'not_authorized' ? 'QuickBooks needs to be reconnected.' :
                                  'QuickBooks is not reachable right now.'

  return (
    <div className="mb-4 rounded-xl border border-amber-200 bg-amber-50 p-4 flex items-start gap-3">
      <span className="text-amber-500 text-xl flex-shrink-0 leading-none">⚠</span>
      <div className="flex-1 text-sm">
        <p className="font-bold text-amber-800">{explain} Invoice generation is unavailable.</p>
        {reason === 'env_missing' && status.missing?.length > 0 && (
          <p className="text-xs text-amber-700 mt-0.5">
            Missing env vars: {status.missing.join(', ')}
          </p>
        )}
        {reason === 'error' && status.error && (
          <p className="text-xs text-amber-700 mt-0.5">{status.error}</p>
        )}
        {reason !== 'env_missing' && (
          <a
            href="/api/qb/auth-start"
            className="inline-block mt-1 text-xs font-bold text-amber-900 hover:underline"
          >
            Click here to reconnect →
          </a>
        )}
      </div>
    </div>
  )
}
