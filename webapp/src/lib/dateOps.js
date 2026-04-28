// Operational-day helpers — mirrors opDay_() in Apps Script (Code.js).
//
// A timestamp's "operational day" is the calendar day a shift would
// be accounted to: for night-shift work that crosses midnight, work
// starting at 02:30 AM on Tuesday belongs to Monday's shift. We pick
// a cutoff hour (default 4 AM) and treat anything below it as the
// PREVIOUS calendar day.
//
// All date math here is local-parts based (year, month, day, hour) so
// DST and timezone shifts don't change a date's bucket.

export const OPERATIONAL_DAY_CUTOFF_HOUR = 4

function isoOf(d) {
  const yyyy = d.getFullYear()
  const mm   = String(d.getMonth() + 1).padStart(2, '0')
  const dd   = String(d.getDate()).padStart(2, '0')
  return `${yyyy}-${mm}-${dd}`
}

/**
 * Bucket a Date into its operational day. Returns YYYY-MM-DD.
 * Hours below cutoffHour roll back to the previous calendar day.
 */
export function opDay(date, cutoffHour = OPERATIONAL_DAY_CUTOFF_HOUR) {
  const d = (date instanceof Date) ? date : new Date(date)
  if (isNaN(d.getTime())) return ''
  const target = (d.getHours() < cutoffHour)
    ? new Date(d.getFullYear(), d.getMonth(), d.getDate() - 1)
    : new Date(d.getFullYear(), d.getMonth(), d.getDate())
  return isoOf(target)
}

/**
 * Combine an ISO calendar date with an HH:MM time-of-day, then
 * opDay-bucket the result. This is the canonical "shift start day"
 * derivation when the user has already entered a Time In on the
 * Sign-In form.
 */
export function opDayFromIsoTime(isoDate, hhmm, cutoffHour = OPERATIONAL_DAY_CUTOFF_HOUR) {
  const md = String(isoDate || '').match(/^(\d{4})-(\d{2})-(\d{2})/)
  const tm = String(hhmm    || '').match(/^(\d{1,2}):(\d{2})/)
  if (!md || !tm) return String(isoDate || '')
  const dt = new Date(
    Number(md[1]), Number(md[2]) - 1, Number(md[3]),
    Number(tm[1]), Number(tm[2])
  )
  return opDay(dt, cutoffHour)
}

/** Today's operational day in the browser's local timezone. */
export function opToday(cutoffHour = OPERATIONAL_DAY_CUTOFF_HOUR) {
  return opDay(new Date(), cutoffHour)
}
