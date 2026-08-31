/**
 * Work Order scanning orchestration service.
 *
 * Handles the flow:
 * 1. User uploads a scan (PDF/image) → stored in R2
 * 2. Scan job enqueued for Python worker
 * 3. Python worker calls Claude Vision → returns parsed WO data
 * 4. This service processes the result → creates WO + marking items
 *
 * The actual Vision parsing stays in Python (parse_work_order.py).
 * This service only orchestrates the API-side steps.
 */
import { eq, and, desc } from 'drizzle-orm';
import type { Db } from '../db/client.js';
import { jobs } from '../db/schema.js';
import { createWorkOrder, getWorkOrderByNumber, updateWorkOrder } from '../db/queries/workOrders.js';
import { createMarkingItemsBulk } from '../db/queries/markingItems.js';
import { createAuditEntry } from '../db/queries/audit.js';

export interface ScanResult {
  workOrderId: string;
  contractor: string;
  contractNum: string;
  regionCode: string;
  location: string;
  fromStreet?: string;
  toStreet?: string;
  dueDate?: string;
  priority?: string;
  workType?: string;
  woReceivedDate?: string;
  waterBlastRequired?: string;
  waterBlastConfirmed?: string;
  waterBlastSqft?: string;
  generalRemarks?: string;
  school?: string;
  prepBy?: string;
  dateEntered?: string;
  topMarkings?: { category: string; description: string }[];
  intersectionGrid?: { intersection: string; n: string; e: string; s: string; w: string; stopMsg: string; stopLines: string }[];
  bikeLaneMarkings?: { type: string; quantity: number | null; source: string }[];
}

/**
 * Process a completed WO scan result from the Python worker.
 * Creates or updates the WO and seeds marking items from scan data.
 */
export async function processScanResult(
  db: Db,
  orgId: string,
  contractorId: string,
  scanResult: ScanResult,
  scanFileKey: string,
  scanCombinedId: string | undefined,
  originalFilename: string,
): Promise<{ woId: string; duplicate: boolean }> {
  // Check for duplicate
  const existing = await getWorkOrderByNumber(db, orgId, scanResult.workOrderId);
  if (existing) {
    return { woId: existing.id, duplicate: true };
  }

  // Create work order
  const wo = await createWorkOrder(db, orgId, {
    orgId,
    woNumber: scanResult.workOrderId,
    contractorId,
    contractNum: scanResult.contractNum || undefined,
    regionCode: scanResult.regionCode || undefined,
    location: scanResult.location || undefined,
    fromStreet: scanResult.fromStreet || undefined,
    toStreet: scanResult.toStreet || undefined,
    dueDate: scanResult.dueDate || undefined,
    priority: scanResult.priority || undefined,
    workType: scanResult.workType || undefined,
    woReceivedDate: scanResult.woReceivedDate || undefined,
    waterBlastRequired: scanResult.waterBlastRequired || undefined,
    waterBlastConfirmed: scanResult.waterBlastConfirmed || undefined,
    waterBlastSqft: scanResult.waterBlastSqft || undefined,
    generalRemarks: scanResult.generalRemarks || undefined,
    school: scanResult.school || undefined,
    prepBy: scanResult.prepBy || undefined,
    dateEntered: scanResult.dateEntered || undefined,
    scanFileKey,
    scanCombinedId,
    originalFilename,
    scanData: {
      topMarkings: scanResult.topMarkings,
      intersectionGrid: scanResult.intersectionGrid,
      bikeLaneMarkings: scanResult.bikeLaneMarkings,
    },
    status: 'received',
  });

  // Seed marking items from scan data
  const markingRows: { category: string; woSection: string; description?: string; intersection?: string; direction?: string; addedBy: string; quantity?: string; unit?: string }[] = [];

  // Top markings → marking items
  if (scanResult.topMarkings) {
    for (const tm of scanResult.topMarkings) {
      if (tm.category && tm.description) {
        markingRows.push({
          category: tm.category,
          woSection: 'top_table',
          description: tm.description,
          addedBy: 'scanner',
        });
      }
    }
  }

  // Intersection grid → marking items (HVX Crosswalk, Stop Msg, Stop Line)
  // Port of Code.js seedMarkingItems_ (lines 10191–10229).
  if (scanResult.intersectionGrid) {
    const { expandDirections } = await import('./markingItemLogic.js');
    for (const ig of scanResult.intersectionGrid) {
      if (!ig.intersection) continue;

      // MI-4: Any non-empty direction cell → HVX Crosswalk item
      // (old app accepted any truthy value, not just "HVX")
      for (const dir of ['n', 'e', 's', 'w'] as const) {
        if (ig[dir] && String(ig[dir]).trim()) {
          markingRows.push({
            category: 'HVX Crosswalk',
            woSection: 'intersection_grid',
            intersection: ig.intersection,
            direction: dir.toUpperCase(),
            addedBy: 'scanner',
          });
        }
      }

      // MI-3: Stop Msg/Lines expanded per-direction
      // e.g. "EW" → two rows: direction E and direction W
      if (ig.stopMsg) {
        const dirs = expandDirections(ig.stopMsg);
        if (dirs.length > 0) {
          for (const d of dirs) {
            markingRows.push({
              category: 'Stop Msg',
              woSection: 'intersection_grid',
              intersection: ig.intersection,
              direction: d,
              addedBy: 'scanner',
            });
          }
        } else {
          // No parseable direction — create one row without direction
          markingRows.push({
            category: 'Stop Msg',
            woSection: 'intersection_grid',
            intersection: ig.intersection,
            addedBy: 'scanner',
          });
        }
      }
      if (ig.stopLines) {
        const dirs = expandDirections(ig.stopLines);
        if (dirs.length > 0) {
          for (const d of dirs) {
            markingRows.push({
              category: 'Stop Line',
              woSection: 'intersection_grid',
              intersection: ig.intersection,
              direction: d,
              addedBy: 'scanner',
            });
          }
        } else {
          markingRows.push({
            category: 'Stop Line',
            woSection: 'intersection_grid',
            intersection: ig.intersection,
            addedBy: 'scanner',
          });
        }
      }
    }
  }

  // Bike lane markings → marking items (preformed symbols from general remarks / bike lane section)
  if (scanResult.bikeLaneMarkings) {
    for (const bm of scanResult.bikeLaneMarkings) {
      if (!bm.type) continue;
      // Map scan type names to canonical marking category names
      const categoryMap: Record<string, string> = {
        'Bike Symbol': 'Old Bike Symbol (w/ rider)',
        'Bike Arrow': 'Bike Lane Arrow',
        'Pedestrian Men': 'Pedestrian Men',
      };
      const category = categoryMap[bm.type] || bm.type;
      markingRows.push({
        category,
        woSection: 'manual',
        addedBy: 'scanner',
        quantity: bm.quantity != null ? String(bm.quantity) : undefined,
        unit: 'EA',
        description: bm.source ? `Source: ${bm.source}` : undefined,
      });
    }
  }

  // WO-5: PT- (Paint/MMA) WOs auto-seed a default Color Surface item.
  // Port of Code.js seedMarkingItems_ (lines 10275–10293).
  // Defaults to 'Bike Lane' (SF) — crew switches to Bus Lane / Pedestrian
  // Space if needed. Color/Material left blank (required before completion).
  const isPaintWO = String(scanResult.workOrderId || '').toUpperCase().startsWith('PT-');
  if (isPaintWO) {
    markingRows.push({
      category: 'Bike Lane',
      woSection: 'top_table',
      addedBy: 'scanner',
      unit: 'SF',
      description: 'Color Surface — confirm type & color',
    });
  }

  if (markingRows.length > 0) {
    await createMarkingItemsBulk(db, orgId, wo.id, markingRows);
  }

  // Auto-geocode the WO location for the map.
  // Port of Code.js geocodeWO_ — builds an address from location + from_street,
  // geocodes it via Google Maps, and saves lat/lng on the WO.
  try {
    const { googleMapsGeocoding } = await import('../integrations/geocoding/googleMaps.js');
    const location = scanResult.location || '';
    const fromStreet = scanResult.fromStreet || '';
    if (location && fromStreet) {
      // NYC borough names for address context
      const boroughNames: Record<string, string> = {
        BK: 'Brooklyn', M: 'Manhattan', BX: 'Bronx', QU: 'Queens', SI: 'Staten Island',
      };
      const borough = boroughNames[scanResult.regionCode || ''] || '';
      const address = `${location} & ${fromStreet}, ${borough}, New York, NY`;
      const result = await googleMapsGeocoding.geocode(address, {
        // Bias to NYC bounds
        bounds: { south: 40.49, west: -74.26, north: 40.92, east: -73.70 },
      });
      if (result) {
        await updateWorkOrder(db, orgId, wo.id, {
          latitude: String(result.lat),
          longitude: String(result.lng),
        });
      } else {
        await updateWorkOrder(db, orgId, wo.id, {
          geocodeWarning: `Could not geocode: ${address}`,
        });
      }
    }
  } catch (err: any) {
    console.warn(`[Scan] Geocoding failed for ${scanResult.workOrderId}:`, err.message);
  }

  return { woId: wo.id, duplicate: false };
}

/**
 * Get the status of pending scan jobs for an org.
 */
export async function getScanJobStatuses(db: Db, orgId: string, jobIds?: string[]) {
  let query = db.select()
    .from(jobs)
    .where(and(
      eq(jobs.orgId, orgId),
      eq(jobs.type, 'scan_work_order'),
    ))
    .orderBy(desc(jobs.createdAt))
    .limit(50);

  return query;
}
