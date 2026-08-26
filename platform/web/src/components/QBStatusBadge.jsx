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
    // Light polling — every 2 minutes catches expirations without
    // hammering the QB CompanyInfo endpoint.
    const id = setInterval(load, 120_000)
    return () => { cancelled = true; clearInterval(id) }
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
