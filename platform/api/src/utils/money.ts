/**
 * Scaled-integer arithmetic for exact decimal math.
 * Matches QuickBooks UnitPrice precision (4 decimal places for rates,
 * 2 decimal places for amounts). Prevents the 1-cent rounding errors
 * that cause QB invoice rejection.
 */

const SCALE_2 = 100;
const SCALE_4 = 10000;

/** Multiply qty × rate with exact 2dp result (for amounts). */
export function money2(qty: number, rate: number): number {
  return Math.round(qty * rate * SCALE_2) / SCALE_2;
}

/** Round a rate to 4 decimal places. */
export function rate4(n: number): number {
  return Math.round(n * SCALE_4) / SCALE_4;
}
