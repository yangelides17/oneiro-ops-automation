// Display helpers for month-end documents (Employee Utilization,
// Certificates).
//
// Shared because the same document has to read identically in four
// places — the approvals queue row, the upload modal's per-page
// dropdown, its "will file" summary, and the enlarged page viewer's
// header. When each built its own string they drifted, and a label that
// omits the month is genuinely ambiguous: the same contract has one of
// these every single month.

// "2026-07" → "Jul 2026". Short form on purpose — it rides on the end of
// an already-long label next to a contract number and a prime.
export function fmtMonthShort(monthIso) {
  const [y, m] = String(monthIso || '').split('-').map(Number)
  if (!y || !m) return monthIso || ''
  return new Date(y, m - 1, 1).toLocaleString(undefined, { month: 'short', year: 'numeric' })
}

// "2026-07" → "July 2026". For standalone use, where there's room.
export function fmtMonthLong(monthIso) {
  const [y, m] = String(monthIso || '').split('-').map(Number)
  if (!y || !m) return monthIso || ''
  return new Date(y, m - 1, 1).toLocaleString(undefined, { month: 'long', year: 'numeric' })
}

// The canonical one-line identity of a month-end document.
// e.g. "Employee Utilization · 84125MBTP701-SI · Delan · Jul 2026"
//
// `withContractor` is off for narrow contexts (the queue row) where the
// prime doesn't fit and the contract number already disambiguates.
export function docLabel(c, { withContractor = true } = {}) {
  if (!c) return ''
  const parts = [c.label, `${c.contract_num}-${c.borough}`]
  if (withContractor && c.contractor) parts.push(c.contractor)
  if (c.month) parts.push(fmtMonthShort(c.month))
  return parts.join(' · ')
}
