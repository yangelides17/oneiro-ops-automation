import { useState } from 'react'

// Optional "Upload Paystub" affordance shown inside the Certified Payroll
// Generate modal. Reads a payrollforconstruction.com paystub (Pre-Check
// Register PDF, or a JPEG/PNG screenshot) via Claude vision and hands the
// per-employee { name, gross_pay, net_pay, deductions } rows back to the
// parent, which forwards them into the CP generation request so the
// Withholdings & Net Pay columns fill automatically.
//
// Props:
//   onParsed(employees | null) — fires with the parsed rows on success,
//                                and null when the user removes the file.

const MAX_BYTES = 20 * 1024 * 1024   // matches the server multer limit
const ACCEPT    = 'application/pdf,image/jpeg,image/png'

const fmtMoney = v => (v == null ? '—' : `$${Number(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`)

export default function PaystubUpload({ onParsed }) {
  const [status, setStatus]       = useState('idle')   // 'idle' | 'parsing' | 'done' | 'error'
  const [error, setError]         = useState('')
  const [employees, setEmployees] = useState([])
  const [filename, setFilename]   = useState('')

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
      setEmployees(rows)
      setFilename(json.filename || file.name)
      setStatus('done')
      onParsed?.(rows)
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
    setEmployees([]); setFilename(''); setStatus('idle'); setError('')
    onParsed?.(null)
  }

  if (status === 'parsing') {
    return (
      <div className="flex items-center gap-2 justify-center py-4 border-2 border-dashed border-slate-200 rounded-xl">
        <div className="w-4 h-4 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
        <span className="text-[12px] text-slate-500">Reading paystub…</span>
      </div>
    )
  }

  if (status === 'done') {
    return (
      <div className="text-left border border-slate-200 rounded-xl p-3 space-y-2">
        <div className="flex items-center justify-between gap-2">
          <p className="text-[12px] font-semibold text-slate-700 truncate">
            ✓ {employees.length} employee{employees.length === 1 ? '' : 's'} read
            <span className="font-normal text-slate-400"> · {filename}</span>
          </p>
          <button
            type="button"
            onClick={remove}
            className="text-[11px] font-bold text-slate-400 hover:text-red-500 flex-shrink-0"
          >
            Remove
          </button>
        </div>
        <div className="max-h-[140px] overflow-y-auto">
          <table className="w-full text-[11px]">
            <thead className="text-slate-400">
              <tr className="text-left">
                <th className="font-medium py-0.5 pr-2">Employee</th>
                <th className="font-medium py-0.5 px-1 text-right">Gross</th>
                <th className="font-medium py-0.5 px-1 text-right">Deduct.</th>
                <th className="font-medium py-0.5 pl-1 text-right">Net</th>
              </tr>
            </thead>
            <tbody className="text-slate-600">
              {employees.map((e, i) => (
                <tr key={i} className="border-t border-slate-100">
                  <td className="py-0.5 pr-2 truncate max-w-[130px]">{e.name}</td>
                  <td className="py-0.5 px-1 text-right tabular-nums">{fmtMoney(e.gross_pay)}</td>
                  <td className="py-0.5 px-1 text-right tabular-nums">{fmtMoney(e.deductions)}</td>
                  <td className="py-0.5 pl-1 text-right tabular-nums">{fmtMoney(e.net_pay)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p className="text-[10px] text-slate-400">
          Withholdings &amp; Net Pay will fill from these values. Gross is cross-checked against recorded hours.
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
          Pre-Check Register — PDF or screenshot. Auto-fills Withholdings &amp; Net Pay.
        </p>
      </label>
      {status === 'error' && error && (
        <p className="text-[11px] text-red-600">{error}</p>
      )}
    </div>
  )
}
