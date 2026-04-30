// Operational Qty thresholds for marking types that the field crew
// frequently mixes up on the form. If the entered value falls outside
// the typical range, the UI raises a confirmation modal so the user
// can verify they meant it (no hard rejection — just a sanity check).
//
// Marking-type strings here MUST match MARKING_CATEGORIES in
// markingCategories.js exactly.

const RULES = {
  'HVX Crosswalk': {
    unit: 'LF',
    inRange: (n) => n > 100,
    message: (n) =>
      `You entered ${formatQty(n)} LF for HVX Crosswalk. ` +
      `Typical crosswalks run more than 100 LF. ` +
      `Are you sure?`,
  },
  'Stop Line': {
    unit: 'LF',
    inRange: (n) => n > 10 && n < 75,
    message: (n) =>
      `You entered ${formatQty(n)} LF for Stop Line. ` +
      `Typical stop lines are between 10 LF and 75 LF. ` +
      `Are you sure?`,
  },
  'Stop Msg': {
    unit: 'EA',
    inRange: (n) => n <= 4,
    message: (n) =>
      `You entered ${formatQty(n)} EA for Stop Msg. ` +
      `Typical stop messages are 4 or fewer per input. ` +
      `Are you sure?`,
  },
}

function formatQty(n) {
  return Number.isInteger(n) ? String(n) : String(n)
}

// Returns { ok: true } when no validation is needed or the value is in
// range. Returns { ok: false, message } when the value is out of range
// — the caller should show a confirmation modal with that message.
//
// Empty / non-numeric input returns ok: true (validation is for
// committed numeric values only).
export function validateQty(category, rawValue) {
  if (rawValue == null || rawValue === '') return { ok: true }
  const n = parseFloat(rawValue)
  if (!isFinite(n)) return { ok: true }
  const rule = RULES[category]
  if (!rule) return { ok: true }
  if (rule.inRange(n)) return { ok: true }
  return { ok: false, message: rule.message(n) }
}
