/**
 * Operational day calculation.
 *
 * Night-shift crews finish work in the early morning hours. A timestamp
 * before the cutoff hour (e.g., 3 AM with a 5 AM cutoff) belongs to the
 * PREVIOUS calendar day for accounting purposes.
 *
 * Ported from Code.js opDay_() and webapp/src/lib/dateOps.js opDay().
 * This is now the SINGLE SOURCE OF TRUTH — no more duplication.
 */

/**
 * Given a Date and a cutoff hour, returns the operational day as YYYY-MM-DD.
 * If the hour is before the cutoff, the date rolls back one day.
 */
export function opDay(date: Date, cutoffHour: number): string {
  const d = new Date(date);
  if (d.getHours() < cutoffHour) {
    d.setDate(d.getDate() - 1);
  }
  return formatDate(d);
}

/**
 * Given an ISO date string (YYYY-MM-DD) and a time string (HH:MM),
 * computes the operational day.
 */
export function opDayFromIsoTime(isoDate: string, hhmm: string, cutoffHour: number): string {
  const [y, m, d] = isoDate.split('-').map(Number);
  const [h] = hhmm.split(':').map(Number);
  const dt = new Date(y, m - 1, d, h);
  return opDay(dt, cutoffHour);
}

/**
 * Returns today's operational day for a given timezone and cutoff.
 */
export function opToday(timezone: string, cutoffHour: number): string {
  const now = new Date(new Date().toLocaleString('en-US', { timeZone: timezone }));
  return opDay(now, cutoffHour);
}

/** Format a Date as YYYY-MM-DD. */
function formatDate(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}
