/**
 * Status Format Utility
 *
 * The DB stores WO status as snake_case ('in_progress', 'received', etc.).
 * The frontend expects Title Case ('In Progress', 'Received', etc.) to
 * match the old Apps Script app's format.
 *
 * This module normalizes at the API boundary so the DB stays clean
 * and the frontend stays unchanged.
 */

const SNAKE_TO_TITLE: Record<string, string> = {
  received:    'Received',
  dispatched:  'Dispatched',
  in_progress: 'In Progress',
  completed:   'Completed',
  returned:    'Returned',
};

const TITLE_TO_SNAKE: Record<string, string> = {
  'Received':    'received',
  'Dispatched':  'dispatched',
  'In Progress': 'in_progress',
  'Completed':   'completed',
  'Returned':    'returned',
};

/** Convert DB status to frontend-facing Title Case. */
export function statusToDisplay(dbStatus: string | null | undefined): string {
  return SNAKE_TO_TITLE[String(dbStatus || '').toLowerCase()] || String(dbStatus || '');
}

/** Convert frontend Title Case to DB snake_case. */
export function statusToDb(displayStatus: string | null | undefined): string {
  const s = String(displayStatus || '').trim();
  return TITLE_TO_SNAKE[s] || s.toLowerCase().replace(/\s+/g, '_');
}
