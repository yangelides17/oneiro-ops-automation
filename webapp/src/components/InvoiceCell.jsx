import { useState } from 'react'

// Invoice column cell. One of: already-invoiced link · em-dash (WO not
// completed) · spinner (in-flight) · error/Retry · "Generate" (disabled
// when QB not connected). The button POSTs /api/qb/invoice/:wo_id and
// propagates the result via onInvoiced so the parent can patch local
// state. Extracted from Dashboard so the WO Tracker and the Doc Status
// docs-queue share the exact same invoice control + endpoint.
export default function InvoiceCell({ wo, qbConnected, onInvoiced }) {
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')

  // 1) Already invoiced — show the doc # as a link to QB
  if (wo.invoice_doc_number) {
    return (
      <a
        href={wo.invoice_view_url || '#'}
        target="_blank"
        rel="noopener noreferrer"
        title={`$${(wo.invoice_amount ?? 0).toLocaleString()} · ${wo.invoice_date || ''}`}
        className="text-xs font-bold px-2 py-1 rounded-lg bg-emerald-100 text-emerald-700 hover:bg-emerald-200"
      >
        #{wo.invoice_doc_number}
      </a>
    )
  }

  // 2) WO isn't Completed — not eligible
  if (String(wo.status || '').toLowerCase() !== 'completed') {
    return <span className="text-slate-300 text-xs px-2 py-1" title="Available once WO is Completed">—</span>
  }

  // 3) In flight
  if (loading) {
    return (
      <span className="inline-flex items-center gap-1 text-xs text-slate-500">
        <span className="w-3 h-3 border-2 border-slate-300 border-t-navy rounded-full animate-spin" />
        Sending…
      </span>
    )
  }

  const doGenerate = async () => {
    setLoading(true)
    setError('')
    try {
      const res = await fetch(`/api/qb/invoice/${encodeURIComponent(wo.id)}`, { method: 'POST' })
      const data = await res.json().catch(() => ({ ok: false, error: 'Non-JSON response' }))
      if (!data.ok) {
        // needs_pricing list → surface a more helpful message
        const detail = Array.isArray(data.needs_pricing) && data.needs_pricing.length > 0
          ? ` (${data.needs_pricing.length} item${data.needs_pricing.length === 1 ? '' : 's'} need pricing)`
          : ''
        throw new Error((data.error || `HTTP ${res.status}`) + detail)
      }
      onInvoiced?.(wo.id, data)
    } catch (e) {
      setError(e.message || String(e))
    } finally {
      setLoading(false)
    }
  }

  // 4) Error — show Retry with the message in a tooltip
  if (error) {
    return (
      <button
        onClick={doGenerate}
        title={error}
        className="text-xs font-bold px-2 py-1 rounded-lg bg-red-100 text-red-700 hover:bg-red-200"
      >
        Retry
      </button>
    )
  }

  // 5) Idle — Generate (disabled when QB disconnected)
  return (
    <button
      onClick={doGenerate}
      disabled={!qbConnected}
      title={qbConnected ? 'Generate QuickBooks invoice' : 'QuickBooks not connected'}
      className={`text-xs font-bold px-2 py-1 rounded-lg ${qbConnected
        ? 'bg-navy text-white hover:bg-navy/80'
        : 'bg-slate-100 text-slate-400 cursor-not-allowed'}`}
    >
      Generate
    </button>
  )
}
