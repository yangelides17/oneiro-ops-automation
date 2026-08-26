/**
 * Production dashboard aggregation service.
 *
 * Computes SF/LF/EA totals by date, contractor, category, and WO
 * from completed marking items within a date range.
 * Also computes shift KPIs (days worked, streak, % utilization).
 */
import { eq, and, between, sql, asc, desc } from 'drizzle-orm';
import type { Db } from '../db/client.js';
import { markingItems, workOrders, contractors } from '../db/schema.js';

export interface ProductionFilters {
  startDate: string;
  endDate: string;
}

export async function getProductionData(db: Db, orgId: string, filters: ProductionFilters) {
  const { startDate, endDate } = filters;

  // Load all completed items with WO + contractor metadata
  const rows = await db.select({
    dateCompleted: markingItems.dateCompleted,
    unit: markingItems.unit,
    quantity: markingItems.quantity,
    category: markingItems.category,
    contractorId: workOrders.contractorId,
    woId: workOrders.id,
    woNumber: workOrders.woNumber,
    location: workOrders.location,
    contractorName: contractors.name,
  })
    .from(markingItems)
    .innerJoin(workOrders, eq(markingItems.woId, workOrders.id))
    .leftJoin(contractors, eq(workOrders.contractorId, contractors.id))
    .where(and(
      eq(markingItems.orgId, orgId),
      eq(markingItems.status, 'completed'),
      between(markingItems.dateCompleted, startDate, endDate),
    ))
    .orderBy(asc(markingItems.dateCompleted));

  // ─── Aggregation ───────────────────────────────────────────
  const dailyMap = new Map<string, { SF: number; LF: number; EA: number; items: number }>();
  const contractorMap = new Map<string, { contractorName: string; SF: number; LF: number; EA: number; items: number }>();
  const categoryMap = new Map<string, { unit: string; qty: number; items: number }>();
  const woMap = new Map<string, { woNumber: string; contractorName: string; location: string; SF: number; LF: number; EA: number; items: number }>();
  let totalSf = 0, totalLf = 0, totalEa = 0, totalItems = 0;

  for (const row of rows) {
    const date = String(row.dateCompleted || '');
    const unit = String(row.unit || '').toUpperCase();
    const qty = Number(row.quantity) || 0;
    if (qty <= 0) continue;

    if (unit === 'SF') totalSf += qty;
    else if (unit === 'LF') totalLf += qty;
    else if (unit === 'EA') totalEa += qty;
    totalItems++;

    // Daily
    const daily = dailyMap.get(date) || { SF: 0, LF: 0, EA: 0, items: 0 };
    if (unit === 'SF') daily.SF += qty;
    else if (unit === 'LF') daily.LF += qty;
    else if (unit === 'EA') daily.EA += qty;
    daily.items++;
    dailyMap.set(date, daily);

    // By contractor
    const cid = row.contractorId;
    const cName = row.contractorName || 'Unknown';
    const c = contractorMap.get(cid) || { contractorName: cName, SF: 0, LF: 0, EA: 0, items: 0 };
    if (unit === 'SF') c.SF += qty;
    else if (unit === 'LF') c.LF += qty;
    else if (unit === 'EA') c.EA += qty;
    c.items++;
    contractorMap.set(cid, c);

    // By category
    const cat = row.category || 'Unknown';
    const catKey = `${cat}|${unit}`;
    const catEntry = categoryMap.get(catKey) || { unit, qty: 0, items: 0 };
    catEntry.qty += qty;
    catEntry.items++;
    categoryMap.set(catKey, catEntry);

    // By WO
    const woEntry = woMap.get(row.woId) || {
      woNumber: row.woNumber,
      contractorName: cName,
      location: row.location || '',
      SF: 0, LF: 0, EA: 0, items: 0,
    };
    if (unit === 'SF') woEntry.SF += qty;
    else if (unit === 'LF') woEntry.LF += qty;
    else if (unit === 'EA') woEntry.EA += qty;
    woEntry.items++;
    woMap.set(row.woId, woEntry);
  }

  // ─── Shift KPIs ───────────────────────────────────────────
  const datesWorked = new Set(dailyMap.keys());
  const daysInRange = daysBetween(startDate, endDate);
  const shiftCount = datesWorked.size;
  // L-12: One decimal place to match old app (e.g. 73.3 not 73)
  const pctDaysWorked = daysInRange > 0 ? Math.round((shiftCount / daysInRange) * 1000) / 10 : 0;

  // Longest consecutive streak
  const sortedDates = [...datesWorked].sort();
  let longestStreak = 0;
  let currentStreak = 0;
  let prevDate = '';
  for (const d of sortedDates) {
    if (prevDate && isNextDay(prevDate, d)) {
      currentStreak++;
    } else {
      currentStreak = 1;
    }
    longestStreak = Math.max(longestStreak, currentStreak);
    prevDate = d;
  }

  // ─── Build response ───────────────────────────────────────
  // by_category: array of { category, unit, qty, items } sorted by qty desc
  const byCategory = Array.from(categoryMap.entries())
    .map(([key, val]) => {
      const [category] = key.split('|');
      return { category, ...val };
    })
    .sort((a, b) => b.qty - a.qty);

  // top_wos: top 25 WOs by total quantity (SF+LF+EA), not item count
  const topWos = Array.from(woMap.values())
    .sort((a, b) => (b.SF + b.LF + b.EA) - (a.SF + a.LF + a.EA))
    .slice(0, 25);

  return {
    range: { start: startDate, end: endDate },
    totals: { SF: totalSf, LF: totalLf, EA: totalEa, items: totalItems },
    shifts: {
      count: shiftCount,
      days_in_range: daysInRange,
      pct_days_worked: pctDaysWorked,
      longest_streak: longestStreak,
    },
    daily: Array.from(dailyMap.entries()).map(([date, d]) => ({ date, ...d })),
    byContractor: Array.from(contractorMap.values()).map(c => ({
      ...c,
      contractor: c.contractorName,  // old-app alias (M-21)
    })),
    by_category: byCategory,
    top_wos: topWos.map(w => ({
      ...w,
      wo_id: w.woNumber,  // old-app alias
    })),
  };
}

function daysBetween(start: string, end: string): number {
  const s = new Date(start);
  const e = new Date(end);
  return Math.max(1, Math.round((e.getTime() - s.getTime()) / (24 * 60 * 60 * 1000)) + 1);
}

// L-13: Use UTC date parsing to avoid DST drift at timezone boundaries
function isNextDay(a: string, b: string): boolean {
  const [ay, am, ad] = a.split('-').map(Number);
  const [by, bm, bd] = b.split('-').map(Number);
  const da = Date.UTC(ay, am - 1, ad);
  const db = Date.UTC(by, bm - 1, bd);
  return db - da === 24 * 60 * 60 * 1000;
}
