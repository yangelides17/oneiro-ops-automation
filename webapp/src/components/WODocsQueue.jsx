import { useMemo, useState } from 'react'
import { Link } from 'react-router-dom'
import DocStatusChips from './DocStatusChips'
import InvoiceCell from './InvoiceCell'
import FilterBar from './FilterBar'

/**
 * WODocsQueue — a condensed "what still needs doc work" list for the top
 * of the Doc Status tab. Shows every COMPLETED work order whose CFR or
 * Invoice isn't yet both Done AND Sent, so the admin sees outstanding
 * doc work in one place.
 *
 * Reuses the WO Tracker's exact pieces (DocStatusChips pills → same
 * /api/documents/flags, InvoiceCell → same /api/qb/invoice/:id) and is
 * fed the SAME `wos` array the WO Tracker renders (lifted from the parent
 * Dashboard). So the two lists are one source of truth: a toggle or an
 * invoice here optimistically updates the shared state and the row drops
 * out the moment its CFR + Invoice are both done + sent.
 */
const isCompleted = (wo) => String(wo?.status || '').toLowerCase() === 'completed'
const docFull = (d) => !!(d && d.done && d.sent)

// Completed WO whose CFR or Invoice isn't fully done+sent.
function isDocsPending(wo) {
  if (!isCompleted(wo)) return false
  return !docFull(wo.docs?.cfr) || !docFull(wo.docs?.invoice)
}

export default function WODocsQueue({ wos, qbConnected, onDocsChange, onInvoiced }) {
  const [contFilt, setContFilt] = useState(() => new Set())
  const toggleCont = (opt) => setContFilt(prev => {
    const next = new Set(prev)
    if (next.has(opt)) next.delete(opt); else next.add(opt)
    return next
  })
  const clearCont = () => setContFilt(new Set())

  const pending = useMemo(() => (wos || []).filter(isDocsPending), [wos])
  const contractors = useMemo(
    () => [...new Set(pending.map(w => w.contractor).filter(Boolean))].sort(),
    [pending]
  )
  const rows = useMemo(
    () => pending.filter(w => contFilt.size === 0 || contFilt.has(w.contractor)),
    [pending, contFilt]
  )

  return (
    <div className="card p-4 space-y-3">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <p className="section-label">
          Completed WOs — Docs Pending{rows.length > 0 ? ` (${rows.length})` : ''}
        </p>
        {contractors.length > 1 && (
          <FilterBar
            label="Contractor"
            options={contractors}
            selected={contFilt}
            onToggle={toggleCont}
            onClear={clearCont}
          />
        )}
      </div>

      {rows.length === 0 ? (
        <p className="text-xs text-slate-400 italic py-2">
          {pending.length === 0
            ? 'All caught up — every completed WO has its CFR and invoice done + sent. ✓'
            : 'No pending WOs for this contractor. ✓'}
        </p>
      ) : (
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-[10px] font-extrabold uppercase tracking-widest text-slate-400 border-b border-slate-100">
                <th className="py-2 px-3">WO #</th>
                <th className="py-2 px-3">Contractor</th>
                <th className="py-2 px-3">Boro</th>
                <th className="py-2 px-3">Location</th>
                <th className="py-2 px-3">Docs</th>
                <th className="py-2 px-3">Invoice</th>
                <th className="py-2 px-3">Drive</th>
              </tr>
            </thead>
            <tbody>
              {rows.map(wo => (
                <tr key={wo.id} className="border-b border-slate-100 hover:bg-slate-50 transition-colors">
                  <td className="py-2.5 px-3 font-mono font-semibold whitespace-nowrap">
                    <Link
                      to={`/field-report?wo=${encodeURIComponent(wo.id)}`}
                      className="text-navy hover:underline"
                    >
                      {wo.id}
                    </Link>
                  </td>
                  <td className="py-2.5 px-3 text-slate-700">{wo.contractor || '—'}</td>
                  <td className="py-2.5 px-3">
                    <span className="bg-slate-100 text-slate-600 rounded px-1.5 py-0.5 text-[11px] font-semibold">
                      {wo.borough}
                    </span>
                  </td>
                  <td className="py-2.5 px-3 text-slate-600 max-w-[180px] truncate">
                    {wo.location || '—'}
                  </td>
                  <td className="py-2.5 px-3">
                    <DocStatusChips woId={wo.id} docs={wo.docs} onChange={onDocsChange} />
                  </td>
                  <td className="py-2.5 px-3">
                    <InvoiceCell wo={wo} qbConnected={qbConnected} onInvoiced={onInvoiced} />
                  </td>
                  <td className="py-2.5 px-3">
                    {wo.folder_url ? (
                      <a
                        href={wo.folder_url}
                        target="_blank"
                        rel="noopener noreferrer"
                        title="Open WO folder in Google Drive"
                        className="text-xs font-bold px-2 py-1 rounded-lg bg-slate-100 text-slate-700 hover:bg-slate-200"
                      >
                        📁
                      </a>
                    ) : (
                      <span className="text-slate-300 text-xs px-2 py-1" title="Folder not yet created">—</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}
