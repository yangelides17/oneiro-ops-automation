import { useState } from 'react'

// Optional "Upload Paystub" affordance shown inside the Certified Payroll
// Generate modal. Reads a payrollforconstruction.com paystub via Claude
// vision and hands the parsed payload back to the parent, which forwards it
// into the CP generation request so the Withholdings & Net Pay columns fill
// automatically.
//
// Two layouts are supported:
//   register — Pre-Check Register: one page, every employee.
//   stubs    — a summary PDF where each page is one employee's check stub.
// Screenshots (JPEG/PNG) of a register work too.
//
// Props:
//   onParsed(payload | null)  — fires with the full parse payload on success,
//                               and null when the user removes the file.
//   expectedWeekStart         — the CP week being generated (YYYY-MM-DD).
//                               Used to flag a wrong-week upload BEFORE the
//                               user clicks Generate; Apps Script blocks it
//                               server-side too, but catching it here saves
//                               a round trip.

const MAX_BYTES = 20 * 1024 * 1024   // matches the server multer limit
const ACCEPT    = 'application/pdf,image/jpeg,image/png'

const fmtMoney = v => (v == null ? '—' : `$${Number(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`)

// 'YYYY-MM-DD' → 'Jul 12'. Parsed by hand so a UTC shift can't move the date.
const fmtDay = (iso) => {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''))
  if (!m) return String(iso || '')
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
  return `${MONTHS[Number(m[2]) - 1]} ${Number(m[3])}`
}

export default function PaystubUpload({ onParsed, expectedWeekStart }) {
  const [status, setStatus]   = useState('idle')   // 'idle' | 'parsing' | 'done' | 'error'
  const [error, setError]     = useState('')
  const [result, setResult]   = useState(null)     // full payload from the parse endpoint

  const parseFile = async (file) => {
    if (!file) return
    if (file.size > MAX_BYTES) {
      setStatus('error')
      setError(`File is too large (${(file.size / 1e6).toFixed(1)} MB). Max is 20 MB.`)
      return
    }
    setStatus('parsing'); setError('')
    try {
      const fd = new FormData()
      fd.append('file', file)
      const res  = await fetch('/api/tools/paystub/parse', { method: 'POST', body: fd })
      const json = await res.json().catch(() => ({}))
      if (!res.ok || json.error) throw new Error(json.error || `HTTP ${res.status}`)
      const rows = Array.isArray(json.employees) ? json.employees : []
      if (rows.length === 0) throw new Error('No employees found on that paystub. Check the file and try again.')
      setResult({ ...json, employees: rows })
      setStatus('done')
      onParsed?.({ ...json, employees: rows })
    } catch (err) {
      setStatus('error')
      setError(err.message || 'Could not read that paystub.')
      onParsed?.(null)
    }
  }

  const onPick = (e) => {
    const f = e.target.files?.[0]
    e.target.value = ''   // allow re-picking the same file
    parseFile(f)
  }

  const remove = () => {
    setResult(null); setStatus('idle'); setError('')
    onParsed?.(null)
  }

  if (status === 'parsing') {
    return (
      <div className="flex flex-col items-center gap-1 py-4 border-2 border-dashed border-slate-200 rounded-xl">
        <div className="flex items-center gap-2">
          <div className="w-4 h-4 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
          <span className="text-[12px] text-slate-500">Reading paystub…</span>
        </div>
        <span className="text-[10px] text-slate-400">Multi-page files take about 15 seconds.</span>
      </div>
    )
  }

  if (status === 'done' && result) {
    const employees = result.employees
    const isStubs   = result.layout === 'stubs'
    const period    = result.pay_period
    const warnings  = Array.isArray(result.warnings) ? result.warnings : []
    // Wrong-week guard. Apps Script blocks this too, but flagging it here
    // means the user never gets as far as pressing Generate.
    const weekMismatch = !!(expectedWeekStart && period?.start && period.start !== expectedWeekStart)
    const anyBonus = employees.some(e => Number(e.bonus) > 0)

    return (
      <div className="text-left border border-slate-200 rounded-xl p-3 space-y-2">
        <div className="flex items-center justify-between gap-2">
          <p className="text-[12px] font-semibold text-slate-700 truncate">
            ✓ {employees.length} employee{employees.length === 1 ? '' : 's'} read
            <span className="font-normal text-slate-400"> · {result.filename}</span>
          </p>
          <button
            type="button"
            onClick={remove}
            className="text-[11px] font-bold text-slate-400 hover:text-red-500 flex-shrink-0"
          >
            Remove
          </button>
        </div>

        <p className="text-[10px] text-slate-400">
          {period ? `Pay period ${fmtDay(period.start)} – ${fmtDay(period.end)}` : 'Pay period not printed'}
          {isStubs && result.page_count != null
            ? ` · ${employees.length} stub${employees.length === 1 ? '' : 's'} across ${result.page_count} page${result.page_count === 1 ? '' : 's'}`
            : ' · Pre-Check Register'}
        </p>

        {weekMismatch && (
          <div className="bg-red-50 border border-red-200 rounded-lg p-2">
            <p className="text-[11px] font-bold text-red-700">
              ⚠ Wrong week — this paystub won’t be accepted
            </p>
            <p className="text-[10px] text-red-600 mt-0.5">
              These stubs cover the week of {fmtDay(period.start)}, but this Certified Payroll is
              the week of {fmtDay(expectedWeekStart)}. Upload the matching week’s paystub.
            </p>
          </div>
        )}

        {warnings.length > 0 && (
          <ul className="bg-amber-50 border border-amber-200 rounded-lg p-2 space-y-0.5 text-[10px] text-amber-800">
            {warnings.map((w, i) => <li key={i}>⚠ {w}</li>)}
          </ul>
        )}

        <div className="max-h-[140px] overflow-y-auto">
          <table className="w-full text-[11px]">
            <thead className="text-slate-400">
              <tr className="text-left">
                <th className="font-medium py-0.5 pr-2">Employee</th>
                <th className="font-medium py-0.5 px-1 text-right">Gross</th>
                {anyBonus && <th className="font-medium py-0.5 px-1 text-right">Bonus</th>}
                <th className="font-medium py-0.5 px-1 text-right">Deduct.</th>
                <th className="font-medium py-0.5 pl-1 text-right">Net</th>
              </tr>
            </thead>
            <tbody className="text-slate-600">
              {employees.map((e, i) => (
                <tr key={i} className="border-t border-slate-100">
                  <td className="py-0.5 pr-2 truncate max-w-[130px]">
                    {isStubs && e.page != null && <span className="text-slate-300 mr-1">p{e.page}</span>}
                    {e.name}
                  </td>
                  <td className="py-0.5 px-1 text-right tabular-nums">{fmtMoney(e.gross_pay)}</td>
                  {anyBonus && (
                    <td className="py-0.5 px-1 text-right tabular-nums text-slate-400">
                      {Number(e.bonus) > 0 ? fmtMoney(e.bonus) : '—'}
                    </td>
                  )}
                  <td className="py-0.5 px-1 text-right tabular-nums">{fmtMoney(e.deductions)}</td>
                  <td className="py-0.5 pl-1 text-right tabular-nums">{fmtMoney(e.net_pay)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p className="text-[10px] text-slate-400">
          Withholdings &amp; Net Pay will fill from these values. Gross is cross-checked against
          recorded hours, with bonus and travel time subtracted first.
        </p>
      </div>
    )
  }

  // idle / error
  return (
    <div className="space-y-1.5 text-left">
      <label className="block border-2 border-dashed border-slate-300 rounded-xl p-4 text-center cursor-pointer hover:border-navy transition-colors">
        <input type="file" accept={ACCEPT} className="hidden" onChange={onPick} />
        <div className="text-2xl mb-0.5">🧾</div>
        <p className="text-[13px] font-semibold text-navy">Upload paystub <span className="font-normal text-slate-400">(optional)</span></p>
        <p className="text-[11px] text-slate-500 mt-0.5">
          Pre-Check Register, or a multi-page PDF with one employee per page.
          Auto-fills Withholdings &amp; Net Pay.
        </p>
      </label>
      {status === 'error' && error && (
        <p className="text-[11px] text-red-600">{error}</p>
      )}
    </div>
  )
}
