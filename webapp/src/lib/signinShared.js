// Shared Sign-In math + constants. Imported by both the Sign-In
// submission page (pages/SignIn.jsx) and the Approvals-page hours editor
// (components/SignInHoursEditor.jsx) so the classification list and the
// hours/OT rules can never drift between where hours are entered and
// where an admin corrects them. Mirror of the server logic in Code.js
// (_allocateDayOvertime_ / _signInRowHours_ / _fmt24to12_).

// Mirrors the dropdown validation on Daily Sign-In Data → Classification.
// Keep in sync with the sheet validation, or submissions fail with
// "Invalid Entry".
export const CLASSIFICATIONS = ['LP', 'SAT']

// Day-of-week-aware hours/OT for a single row. Sat/Sun → all OT; weekday
// over 8h → OT. `workDateIso` is the SHIFT START date (YYYY-MM-DD),
// constructed from local parts so it doesn't UTC-shift the day-of-week.
//
// Cross-midnight: if Time Out <= Time In we assume the shift rolled over
// to the next calendar day and add 24h. Hours stay bucketed under the
// start day for OT — a Fri-night → Sat-morning shift counts as Friday.
export const calcHours = (tin, tout, workDateIso) => {
  if (!tin || !tout) return { hours: '', overtime: '' }
  const [ih, im] = tin.split(':').map(Number)
  const [oh, om] = tout.split(':').map(Number)
  let mins = (oh * 60 + om) - (ih * 60 + im)
  if (mins <= 0) mins += 24 * 60
  const hrs = mins / 60

  const ot = isWeekendIso(workDateIso) ? hrs : Math.max(0, hrs - 8)
  return { hours: hrs.toFixed(2), overtime: ot.toFixed(2) }
}

// True when the YYYY-MM-DD date falls on Saturday or Sunday.
export const isWeekendIso = (dateIso) => {
  const m = String(dateIso || '').match(/^(\d{4})-(\d{2})-(\d{2})/)
  if (!m) return false
  const dow = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getDay()
  return dow === 0 || dow === 6
}

// Apply the Mon–Fri-over-8 / Sat-Sun-all-OT rule to a day total.
export const splitStOt = (hours, dateIso) => {
  if (isWeekendIso(dateIso)) return { st: 0, ot: hours }
  if (hours <= 8) return { st: hours, ot: 0 }
  return { st: 8, ot: hours - 8 }
}

// "7:00 AM" / "13:05" → "07:00" / "13:05" for <input type="time">.
// Accepts the 12-hour strings Daily Sign-In Data stores as well as
// bare 24-hour strings. Returns '' for unparseable input.
export const to24h = (s) => {
  const t = String(s || '').trim()
  if (!t) return ''
  let m = t.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i)
  if (m) {
    let h = Number(m[1]) % 12
    if (m[3].toUpperCase() === 'PM') h += 12
    return `${String(h).padStart(2, '0')}:${m[2]}`
  }
  m = t.match(/^(\d{1,2}):(\d{2})$/)
  if (m) return `${String(Number(m[1])).padStart(2, '0')}:${m[2]}`
  return ''
}
