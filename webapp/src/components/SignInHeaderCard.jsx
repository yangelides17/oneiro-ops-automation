import { useEffect, useState } from 'react'

// Approvals-page replica of the Sign-In tab's header card. Read-only
// reference so an admin can hand-fill a printed sheet's header when the
// crew left it blank — same fields, layout, and formatting as the crew
// see on the Sign-In tab (SignIn.jsx header card), minus the editable
// shift date (the date is fixed by the submitted sheet here).
//
// Self-fetches its metadata from /signin-header (mirrors how
// SignInHoursEditor loads its rows) so it stays independent of the hours
// editor and renders regardless of the hours state.
//
// Props:
//   fileId, filename — the pending sign-in being reviewed.

// Pretty date for the header: "Apr 26 (Sat)". Copied from SignIn.jsx so
// the two cards format the shift date identically.
const prettyQueueDate = (iso) => {
  const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/)
  if (!m) return String(iso)
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]))
  return d.toLocaleDateString(undefined, {
    month: 'short', day: 'numeric', weekday: 'short',
  })
}

// Numerical MM/DD/YY (zero-padded) — the format written on the sheet.
const numericDate = (iso) => {
  const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/)
  return m ? `${m[2]}/${m[3]}/${m[1].slice(2)}` : ''
}

export default function SignInHeaderCard({ fileId, filename }) {
  const [loading, setLoading] = useState(true)
  const [error,   setError]   = useState('')
  const [header,  setHeader]  = useState(null)

  useEffect(() => {
    let cancelled = false
    setLoading(true); setError(''); setHeader(null)
    const qs = new URLSearchParams({ filename: filename || '' }).toString()
    fetch(`/api/approvals/${encodeURIComponent(fileId)}/signin-header?${qs}`)
      .then(r => r.json())
      .then(json => {
        if (cancelled) return
        if (json.error) { setError(json.error); return }
        setHeader(json.header || null)
      })
      .catch(err => { if (!cancelled) setError(err.message || 'Failed to load header') })
      .finally(() => { if (!cancelled) setLoading(false) })
    return () => { cancelled = true }
  }, [fileId, filename])

  if (loading) {
    return (
      <div className="mb-3 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-500">
        Loading sign-in header…
      </div>
    )
  }
  if (error) {
    return (
      <div className="mb-3 rounded-lg border border-amber-200 bg-amber-50 px-3 py-2 text-xs text-amber-800">
        Couldn’t load the sign-in header: {error}
      </div>
    )
  }
  if (!header) return null

  const effectiveDate = header.date || ''
  const wos = Array.isArray(header.wos) ? header.wos : []

  // Markup mirrors the Sign-In tab header card (SignIn.jsx) exactly —
  // same .card / field-label classes and responsive grid — so the admin
  // sees the identical header they'd hand-write. Wrapped in a labeled
  // shell to make clear it's the sheet header (vs. the recorded hours).
  return (
    <div className="mb-3">
      <p className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-1.5">
        Sign-In Sheet Header
      </p>
      <div className="card p-4 space-y-2">
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-sm">
          <div>
            <span className="field-label">Shift Date</span>
            <p className="font-semibold text-slate-400">
              {prettyQueueDate(effectiveDate)}
              {numericDate(effectiveDate) && ` (${numericDate(effectiveDate)})`}
            </p>
          </div>
          <div>
            <span className="field-label">Contract</span>
            <p className="font-semibold text-navy">
              {header.bill_contract_number || header.contract_number} · {header.bill_borough || header.borough}
            </p>
          </div>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <span className="field-label">Prime Contractor</span>
            <p className="text-sm font-semibold text-navy">{header.prime_contractor || header.contractor || '—'}</p>
          </div>
          <div>
            <span className="field-label">Subcontractor</span>
            <p className="text-sm font-semibold text-navy">{header.subcontractor || '—'}</p>
          </div>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <span className="field-label">Address</span>
            <p className="text-sm font-semibold text-navy">{header.address || '—'}</p>
          </div>
          <div>
            <span className="field-label">Agency</span>
            <p className="text-sm font-semibold text-navy">DOT</p>
          </div>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <span className="field-label">Work Orders ({wos.length})</span>
            <ul className="text-sm text-slate-700 space-y-0.5 mt-0.5">
              {wos.map(w => (
                <li key={w.id} className="flex gap-2">
                  <span className="font-mono font-semibold">{w.id}</span>
                  {w.location && (
                    <span className="text-slate-500">| {w.location}</span>
                  )}
                </li>
              ))}
            </ul>
          </div>
          {header.crew_chief && (
            <div>
              <span className="field-label">Crew Chief</span>
              <p className="text-sm font-semibold text-navy">{header.crew_chief}</p>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
