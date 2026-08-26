/**
 * API Compatibility Layer
 *
 * Transforms the new Drizzle schema shapes into the field names
 * the existing React components expect from the old Apps Script API.
 *
 * This is the systematic fix: instead of changing every React component,
 * we transform at the API boundary. Every route that returns WO/marking/signin
 * data should use these transforms.
 */

/** Capitalize status to match old app: "in_progress" → "In Progress" */
export function capitalizeStatus(status: string): string {
  const map: Record<string, string> = {
    received: 'Received',
    dispatched: 'Dispatched',
    in_progress: 'In Progress',
    completed: 'Completed',
    returned: 'Returned',
  };
  return map[status] || status;
}

/**
 * Transform a WO row (from Drizzle) into the shape React components expect.
 * Used by: Dashboard, FieldReport, SignIn, NavTab, ScanWO
 */
export function transformWo(
  wo: Record<string, any>,
  contractorName: string,
  markingRollup?: { total: number; completed: number; qty: number; unit: string },
) {
  return {
    // React uses wo.id as the display WO number (not UUID)
    id: wo.woNumber,
    _uuid: wo.id || wo.uuid,

    // Contractor name (React expects a string, not a UUID)
    contractor: contractorName,
    contractor_id: wo.contractorId,

    // Field name mapping (React uses snake_case)
    contract_num: wo.contractNum || '',
    borough: wo.regionCode || '',
    contract_id: wo.contractId || '',
    location: wo.location || '',
    from_street: wo.fromStreet || '',
    to_street: wo.toStreet || '',
    due_date: wo.dueDate || '',
    priority: wo.priority || '',
    work_type: wo.workType || '',
    wo_received: wo.woReceivedDate || '',
    water_blast: wo.waterBlastRequired || '',
    water_blast_confirmed: wo.waterBlastConfirmed || '',
    water_blast_sqft: wo.waterBlastSqft || '',
    status: capitalizeStatus(wo.status || 'received'),
    dispatch_date: wo.dispatchDate || '',
    work_start: wo.workStartDate || '',
    work_end: wo.workEndDate || '',
    issues: wo.issuesReported || '',
    notes: wo.notes || '',
    general_remarks: wo.generalRemarks || '',
    school: wo.school || '',
    prep_by: wo.prepBy || '',
    date_entered: wo.dateEntered || '',

    // Geo
    latitude: wo.latitude ? Number(wo.latitude) : null,
    longitude: wo.longitude ? Number(wo.longitude) : null,
    geocode_warning: wo.geocodeWarning || '',

    // Marking rollups
    quantity: markingRollup ? String(Math.round(markingRollup.qty)) : (wo.quantity || ''),
    quantity_unit: markingRollup?.unit || wo.quantityUnit || '',
    markings_total: markingRollup?.total ?? wo.markingsTotal ?? 0,
    markings_completed: markingRollup?.completed ?? wo.markingsCompleted ?? 0,

    // Scan
    scan_file_id: wo.scanFileKey || '',
    original_filename: wo.originalFilename || '',

    // Doc flags placeholder (populated separately when needed)
    docs: wo.docs || {
      cfr: { done: false, sent: false },
      production_log: { done: false, sent: false },
      signin: { done: false, sent: false },
      certified_payroll: { done: false, sent: false },
      invoice: { done: false, sent: false },
    },

    folder_url: '', // No Drive folder in new system
  };
}

/**
 * Transform a marking item row into the shape React components expect.
 * Used by: FieldReport
 */
export function transformMarkingItem(item: Record<string, any>) {
  return {
    id: item.id,
    wo_id: item.woId,
    work_type: item.workType || '',
    wo_section: item.woSection || '',
    category: item.category || '',
    intersection: item.intersection || '',
    direction: item.direction || '',
    description: item.description || '',
    quantity: item.quantity ? Number(item.quantity) : null,
    unit: item.unit || '',
    color_material: item.colorMaterial || '',
    date_completed: item.dateCompleted || '',
    status: capitalizeStatus(item.status || 'pending'),
    crew_chief: item.crewChief || '',
    added_by: item.addedBy || '',
    notes: item.notes || '',
  };
}

/**
 * Transform a sign-in entry into the shape React components expect.
 */
export function transformSigninEntry(entry: Record<string, any>, contractorName: string) {
  return {
    id: entry.id,
    date: entry.workDate || '',
    wo_id: entry.woId,
    contractor: contractorName,
    contract_num: entry.contractNum || '',
    borough: entry.regionCode || '',
    location: entry.location || '',
    employee_name: entry.employeeName || '',
    classification: entry.classification || '',
    time_in: entry.timeIn || '',
    time_out: entry.timeOut || '',
    hours: entry.hoursWorked ? Number(entry.hoursWorked) : null,
    overtime: entry.otHours ? Number(entry.otHours) : null,
    crew_chief: entry.crewChief || '',
    admin_reviewed: entry.adminReviewed || false,
    review_notes: entry.reviewNotes || '',
  };
}

/**
 * Transform an employee for the dropdown.
 */
export function transformEmployee(emp: Record<string, any>) {
  return {
    name: emp.name,
    id: emp.id,
  };
}
