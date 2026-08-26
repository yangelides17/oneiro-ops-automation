/**
 * Overtime calculation service — SINGLE SOURCE OF TRUTH.
 *
 * Ported from Code.js _allocateDayOvertime_ (lines 6154–6171)
 * and webapp/src/lib/signinShared.js calcHours().
 *
 * Rules (configurable per org via overtime_rules table):
 * - Daily threshold: first N hours are ST, rest is OT (default 8)
 * - Weekend rule: if shift date is Sat/Sun, ALL hours are OT
 * - Cross-group lookback: 8h cap shared across all sign-in groups
 *   for the same employee on the same day
 * - Cross-midnight: if timeOut <= timeIn, assume shift rolled over
 *   and add 24 hours. Hours bucket under the START date.
 */

export interface OvertimeRules {
  dailyThresholdHours: number | null;  // null = no daily threshold
  weeklyThresholdHours: number | null; // null = no weekly threshold (not used in current system)
  weekendAllOt: boolean;
  crossGroupLookback: boolean;
}

export interface OtEntry {
  key: string;   // normalized employee name
  hours: number; // total hours for this row
}

export interface OtResult {
  hours: number;
  overtime: number;
}

const DEFAULT_RULES: OvertimeRules = {
  dailyThresholdHours: 8,
  weeklyThresholdHours: null,
  weekendAllOt: true,
  crossGroupLookback: true,
};

/**
 * Core OT allocator — exact port of Code.js _allocateDayOvertime_.
 *
 * For each entry, computes how much of that entry's hours are OT,
 * considering hours already worked that day (via `prior` map).
 *
 * The `prior` map is keyed by normalized employee name and contains
 * hours already worked earlier that day (from prior sign-in groups).
 * This enables cross-group lookback.
 *
 * @returns Array of OT hours, aligned 1:1 with entries
 */
export function allocateDayOvertime(
  entries: OtEntry[],
  isWeekend: boolean,
  prior: Record<string, number>,
  rules: OvertimeRules = DEFAULT_RULES,
): number[] {
  const counted: Record<string, number> = { ...prior };

  return entries.map((e) => {
    const hours = e.hours || 0;
    const key = e.key;
    let ot: number;

    if (rules.weekendAllOt && isWeekend) {
      // Weekend: all hours are OT
      ot = hours;
    } else if (rules.dailyThresholdHours !== null) {
      // Weekday with daily threshold: first N hours ST, rest OT
      const threshold = rules.dailyThresholdHours;
      const p = counted[key] || 0;
      const combinedST = Math.min(p + hours, threshold);
      const priorST = Math.min(p, threshold);
      ot = Math.max(0, hours - (combinedST - priorST));
    } else {
      // No daily threshold configured: no OT
      ot = 0;
    }

    counted[key] = (counted[key] || 0) + hours;
    return round2(ot);
  });
}

/**
 * Compute hours worked from time-in and time-out strings.
 * Exact port of Code.js _signInRowHours_ (lines 6194–6201).
 *
 * Handles cross-midnight shifts: if timeOut <= timeIn, adds 24 hours.
 * Returns 0 if either time is unparseable.
 */
export function computeRowHours(timeIn: string, timeOut: string): number {
  const a = parseTimeOfDay(timeIn);
  const b = parseTimeOfDay(timeOut);
  if (!a || !b) return 0;
  let mins = (b.hours * 60 + b.minutes) - (a.hours * 60 + a.minutes);
  if (mins <= 0) mins += 24 * 60;
  return round2(mins / 60);
}

/**
 * Simple ST/OT split for a known total.
 * Port of webapp signinShared.js splitStOt().
 */
export function splitStOt(
  hours: number,
  isWeekend: boolean,
  rules: OvertimeRules = DEFAULT_RULES,
): { st: number; ot: number } {
  if (rules.weekendAllOt && isWeekend) return { st: 0, ot: hours };
  const threshold = rules.dailyThresholdHours ?? Infinity;
  if (hours <= threshold) return { st: hours, ot: 0 };
  return { st: threshold, ot: round2(hours - threshold) };
}

/**
 * Check if a YYYY-MM-DD date string falls on Saturday or Sunday.
 * Port of Code.js _isWeekendDateStr_ (lines 6185–6190).
 */
export function isWeekendDate(dateIso: string): boolean {
  const m = String(dateIso || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return false;
  const dow = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getDay();
  return dow === 0 || dow === 6;
}

/**
 * Parse a time string into 24-hour hours and minutes.
 * Exact port of Code.js _parseSignInTimeOfDay_ (lines 6090–6102).
 *
 * Accepts:
 * - "7:00 AM", "11:30 PM" (12-hour format)
 * - "07:00", "23:30" (24-hour format)
 */
export function parseTimeOfDay(s: string): { hours: number; minutes: number } | null {
  const t = String(s || '').trim();
  if (!t) return null;

  // 12-hour format: "h:mm AM/PM"
  let m = t.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (m) {
    let h = Number(m[1]) % 12;
    if (m[3].toUpperCase() === 'PM') h += 12;
    return { hours: h, minutes: Number(m[2]) };
  }

  // 24-hour format: "HH:MM"
  m = t.match(/^(\d{1,2}):(\d{2})$/);
  if (m) return { hours: Number(m[1]), minutes: Number(m[2]) };

  return null;
}

/**
 * Convert 24-hour time to 12-hour AM/PM format.
 * Port of Code.js _fmt24to12_.
 */
export function to12h(time24: string): string {
  const parsed = parseTimeOfDay(time24);
  if (!parsed) return '';
  const h = parsed.hours;
  const m = String(parsed.minutes).padStart(2, '0');
  if (h === 0) return `12:${m} AM`;
  if (h < 12) return `${h}:${m} AM`;
  if (h === 12) return `12:${m} PM`;
  return `${h - 12}:${m} PM`;
}

/**
 * Convert any time format to 24-hour "HH:MM".
 * Port of webapp signinShared.js to24h().
 */
export function to24h(s: string): string {
  const parsed = parseTimeOfDay(s);
  if (!parsed) return '';
  return `${String(parsed.hours).padStart(2, '0')}:${String(parsed.minutes).padStart(2, '0')}`;
}

/**
 * Normalize an employee name for cross-group matching.
 * Lowercase, trimmed, collapsed whitespace.
 */
export function normalizeEmployeeName(name: string): string {
  return String(name || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

/** Round to 2 decimal places. */
function round2(n: number): number {
  return Math.round(n * 100) / 100;
}
