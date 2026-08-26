/**
 * Work Order Vision Parser — calls Claude Vision to extract WO data from scans.
 * Direct port of platform/worker/parse_work_order.py.
 */
import { config } from '../config.js';

// ── Borough normalization ──────────────────────────────────────
const BOROUGH_MAP: Record<string, string> = {
  K: 'BK', BK: 'BK',
  M: 'M', MN: 'M',
  BX: 'BX',
  Q: 'QU', QU: 'QU',
  SI: 'SI', R: 'SI',
};

const MMA_REMARK_KEYWORDS = [
  'bike lane', 'bus lane', 'pedestrian space', 'pedestrian stop',
  'ped space', 'ped stop', 'mma', 'color surface',
];

// ── Extraction prompt (exact copy from parse_work_order.py) ────
const EXTRACTION_PROMPT = `You are parsing a scanned NYC DOT Pavement Marking Work Order Report form.
The form is issued by NYC DOT and given to pavement marking contractors like Metro Express.

Extract every field listed below and return ONLY valid JSON — no explanation, no markdown fences.
If a field is not visible, illegible, or not applicable, use null. Never guess.

{
  "work_order_id":      "The Work Order number, e.g. PT-11930. Usually in top-left area labeled 'Work Order'.",
  "contractor":         "Prime contractor name, e.g. METRO EXPRESS. Labeled 'Contractor'.",
  "contract_number":    "Full contract number including suffix, e.g. 84122MBTP496/SFT. Labeled 'Contract Number'.",
  "borough":            "Single-letter or abbreviated boro code exactly as printed: K, M, BX, Q or QU, SI, R. Labeled 'Boro'.",
  "location":           "Street name in the 'Location' field.",
  "from_street":        "Street name in the 'From/At' field.",
  "to_street":          "Street name in the 'To' field.",
  "due_date":           "Due Date as printed on the form, e.g. 8/28/2025.",
  "priority_level":     "Priority level label, e.g. '3 - Schedule'. Labeled 'Priority Level'.",
  "pavement_work_type": "Pavement work type, e.g. REFURBISHMENT or NEW. Labeled 'Pavement Work'.",
  "wo_received_date":   "The 'Issue To Contractor Date' at the bottom of the form, e.g. 07/17/2025.",
  "date_entered":       "The 'Date Entered Into Database' (or 'Date Entered') date printed on the form header. If you cannot find an explicit Date-Entered field, use null.",
  "school":             "The 'School' field in the header row. Usually 'NA' if no school is involved; otherwise the school name. Transcribe exactly as printed (often handwritten).",
  "prep_by":            "The 'Prep By' or 'Prepared By' name printed in the bottom block of the form (not a signature — the printed name). Transcribe exactly.",
  "water_blast_sqft":   "If any waterblasting square footage is handwritten anywhere on the form, extract the number as an integer. Otherwise null.",
  "general_remarks":    "The full text of the General Remarks section (middle of the form, labeled 'General Remarks>>>>>'). Transcribe exactly as written, including handwritten text.",

  "top_markings": [
    {
      "category":    "One of: Double Yellow Line, Lane Lines, Gores, Messages, Arrows, Solid Lines, Rail Road X/Diamond, Others. Use the label as printed in the middle of the top table (a row per line/marking type).",
      "description": "The free-text description to the right of the category label, e.g. 'RECAP FROM HAMILTON PL TO 2ND AV'. Transcribe exactly, including 'RECAP' if present. Leave out any rows where the description column is blank."
    }
  ],

  "intersection_grid": [
    {
      "intersection": "The intersection name as printed in the leftmost INTERSECTIONS column, e.g. '5 AV', 'HAMILTON PL'. Transcribe exactly.",
      "n":          "The value of the 'North' column for this row, usually 'HVX' or blank.",
      "e":          "The value of the 'East' column for this row.",
      "s":          "The value of the 'South' column for this row.",
      "w":          "The value of the 'West' column for this row.",
      "stop_msg":   "The value of the 'Stop Msg' column for this row. Usually blank, or a directional string like 'West', 'East', 'EW', 'NSEW'.",
      "stop_lines": "The value of the 'Stop lines' (far-right) column for this row. Same format as stop_msg."
    }
  ],

  "bike_lane_markings": [
    {
      "type":     "Exactly one of: 'Bike Symbol', 'Bike Arrow', 'Pedestrian Men'. See the bike_lane_markings rules below.",
      "quantity": "The integer written immediately before the symbol's abbreviation/name (e.g. 3 from '3 BS'). Use null if a symbol is named with no count.",
      "source":   "'general_remarks' if found in the General Remarks line, or 'bike_lane_section' if found in the 'Bike Lane Work (NEW)' row."
    }
  ]
}

Important notes:
- water_blast_sqft may appear as a handwritten number near the words 'waterblast', 'WB', or 'water blast' anywhere on the form.
- The General Remarks section is critical — it often contains handwritten notes about the type of work (e.g. 'RECAP PAINT FOR BIKE LANE', 'BUS LANE', 'PED SPACE'). Transcribe it fully.
- The Issue To Contractor Date is near the bottom of the form.

bike_lane_markings rules — a SEPARATE, self-contained extraction. It does NOT change
anything above: general_remarks is still transcribed in full, and top_markings /
intersection_grid are unchanged. While you read the General Remarks, ALSO pull out any
counted preform bike-lane / pedestrian symbols into bike_lane_markings. Scan BOTH of
these places:
  (A) The General Remarks line (free text).
  (B) The 'Bike Lane Work (NEW):' cell — the last row of the top table (often
      highlighted yellow), labeled 'Bike Lane Work (NEW)'. If it has text, parse the
      same tokens from it.

Recognize BOTH the abbreviation AND the full written-out name (singular or plural):
  • 'BS'                     or 'Bike Symbol' / 'Bike Symbols'                  → type 'Bike Symbol'
  • 'BA' or 'BSA'            or 'Bike Arrow(s)' / 'Bike Symbol Arrow(s)'        → type 'Bike Arrow'
  • 'PED MEN' / 'PED MAN'    or 'Pedestrian Men' / 'Pedestrian Man'            → type 'Pedestrian Men'
The quantity is the number written immediately before the abbreviation/name. Set
source='bike_lane_section' for tokens from (B), otherwise 'general_remarks'.

Rules:
  • Emit an entry whenever the remarks indicate a bike/ped symbol type needs to be done —
    WITH OR WITHOUT a number. If a number is written immediately before the
    abbreviation/name, set quantity to it; otherwise set quantity to null. A bare mention
    with no count (e.g. 'REFURB BS') still means the symbol is required — emit it with
    quantity null.
  • Emit each type AT MOST ONCE. If a type appears both with and without a number
    (e.g. 'REFURB BS ... 9 BS'), emit it a single time using the number (quantity 9).
  • IGNORE drawing/plan references entirely: tokens like 'MD-762_4', 'MD-882-2,1',
    'MD-19232_3', 'SEE DWG ...', 'DWG ...' are plan numbers, never markings. A digit that
    is part of a drawing number is never a quantity. The letters BS / BA / BSA / PED MEN
    (or the full names) must actually appear — never infer a symbol from a drawing number.
  • Do NOT source anything for this field from the standard top-table rows (Double Yellow /
    Lane Lines / Gores / Messages / Arrows / Solid Lines / Rail Road / Others) or the
    intersection grid — those are handled by top_markings / intersection_grid.
  • If neither place mentions any bike/ped symbol, return an empty array [].

Examples:
  'SEE DWG MD-762_4  3 BS  2 PED MEN  2 BA' →
    [{"type":"Bike Symbol","quantity":3,"source":"general_remarks"},
     {"type":"Pedestrian Men","quantity":2,"source":"general_remarks"},
     {"type":"Bike Arrow","quantity":2,"source":"general_remarks"}]
  'REFURB BS SEE MD-882-2,1  9 BS  6 BSA' →
    [{"type":"Bike Symbol","quantity":9,"source":"general_remarks"},
     {"type":"Bike Arrow","quantity":6,"source":"general_remarks"}]
    ('REFURB BS' and '9 BS' refer to the same Bike Symbols → one entry, quantity 9)
  'REFURB BS SEE DWG MD-1226' →
    [{"type":"Bike Symbol","quantity":null,"source":"general_remarks"}]
    (no count given, but Bike Symbols are still required)
  'SEE DWG MD-19232_3  13 BS' →
    [{"type":"Bike Symbol","quantity":13,"source":"general_remarks"}]

top_markings rules:
- This is the upper table that lists marking CATEGORIES down the middle column (Double Yellow CenterLine / Lane Lines / Gores / Messages / Arrows / Solid Lines / Rail Road X / Diamond / Others).
- Only include a row if its description column contains any text (e.g. 'RECAP', 'RECAP FROM HAMILTON PL TO 2ND AV'). Skip blank rows entirely.
- Normalize the category label: output 'Double Yellow Line' (not 'Double Yellow CenterLine'), 'Rail Road X/Diamond' (not 'Rail Road X / Diamond').
- Preserve the description text verbatim including all caps.
- Order top_markings by their printed row order.

intersection_grid rules:
- This is the bottom table with column headers: INTERSECTIONS | Order | North | East | South | West | Stop Msg | Sch M 8' | Sch M 10' | Stop lines
- IGNORE the Order, Sch M 8', and Sch M 10' columns — they are unused in our system.
- Only include a row if the INTERSECTIONS cell has text AND at least one of N/E/S/W/stop_msg/stop_lines has a non-empty value. Skip blank rows at the bottom of the table.
- For N/E/S/W cells: copy the value verbatim. Typical value is 'HVX' when a crosswalk is required at that direction; blank otherwise.
- For stop_msg and stop_lines cells: copy the directional string verbatim. Values are usually 'North', 'East', 'South', 'West', or concatenations like 'EW' (East AND West), 'NS', 'NSEW'. Preserve the exact letters.
- Return an empty array if the form has no intersection grid entries.
- Order intersection_grid top-to-bottom as printed.`;

// ── Types ──────────────────────────────────────────────────────
export interface ParsedWoData {
  workOrderId: string;
  contractor: string;
  contractNumber: string;
  regionCode: string;
  location: string;
  fromStreet: string;
  toStreet: string;
  dueDate: string;
  priority: string;
  pavementWorkType: string;
  woReceivedDate: string;
  dateEntered: string;
  school: string;
  prepBy: string;
  waterBlastRequired: string;
  waterBlastConfirmed: string;
  waterBlastSqft: string;
  workType: string;
  generalRemarks: string;
  topMarkings: { category: string; description: string }[];
  intersectionGrid: { intersection: string; n: string; e: string; s: string; w: string; stopMsg: string; stopLines: string }[];
  bikeLaneMarkings: { type: string; quantity: number | null; source: string }[];
}

// ── Call Claude Vision ─────────────────────────────────────────
export async function parseWorkOrderScan(fileBytes: Buffer, mimeType: string): Promise<ParsedWoData> {
  if (!config.anthropic.apiKey) {
    throw new Error('ANTHROPIC_API_KEY not configured');
  }

  const encoded = fileBytes.toString('base64');

  const contentBlock = mimeType === 'application/pdf'
    ? { type: 'document' as const, source: { type: 'base64' as const, media_type: mimeType, data: encoded } }
    : { type: 'image' as const, source: { type: 'base64' as const, media_type: mimeType, data: encoded } };

  // Call Claude API directly via fetch (no SDK dependency needed)
  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': config.anthropic.apiKey,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify({
      model: 'claude-opus-4-6',
      max_tokens: 4096,
      messages: [{
        role: 'user',
        content: [
          contentBlock,
          { type: 'text', text: EXTRACTION_PROMPT },
        ],
      }],
    }),
  });

  if (!response.ok) {
    const err = await response.text();
    throw new Error(`Claude API error (${response.status}): ${err}`);
  }

  const result = await response.json() as any;

  // Log actual token usage for cost tracking
  if (result.usage) {
    const u = result.usage;
    const inputCost = (u.input_tokens / 1_000_000) * 15;
    const outputCost = (u.output_tokens / 1_000_000) * 75;
    console.log(`[Scan Cost] input=${u.input_tokens} output=${u.output_tokens} cost=$${(inputCost + outputCost).toFixed(4)}`);
  }

  let raw = result.content?.[0]?.text?.trim() || '';

  // Strip markdown code fences if present
  if (raw.startsWith('```')) {
    const lines = raw.split('\n');
    raw = lines.slice(1, lines[lines.length - 1] === '```' ? -1 : undefined).join('\n');
  }

  const extracted = JSON.parse(raw);
  return normalizeWoData(extracted);
}

// ── Detect WO prompt (exact copy from parse_work_order.py) ─────
const DETECT_WO_PROMPT = `You are examining a multi-page PDF scanned from a stack of NYC DOT Pavement Marking Work Orders.

The stack may contain:
  • Work Order pages (WO) — the NYC DOT Pavement Marking Work Order Report form
  • Contractor Field Report pages (CFR) — the form that's usually stapled to each WO
  • Plan / drawing pages — engineering site diagrams or plan sheets, sometimes
    attached to a WO to clarify the work that needs to be done
  • Blank / separator pages

Your job: identify each Work Order document and return the page numbers that
make up its scan. Rules:

1. A Work Order document always contains exactly ONE WO page.
2. If the page IMMEDIATELY after a WO page is a CFR page, perform this
   strict check before bundling it:
     a. Read the WO # printed on the CFR page.
     b. Read the WO # printed on the preceding WO page.
     c. Compare them CHARACTER-FOR-CHARACTER. Example: "RM-43283" bundles
        with "RM-43283" but NOT with "RM-43285", "RM43283", or "RM-43283A".
     d. Include the CFR in the wo_document ONLY when the two WO #s are
        identical strings.
3. If the next page is a CFR but ANY of the following is true, DO NOT
   include it — the WO stands alone:
     • the CFR's WO # differs from the preceding WO's WO # (even by one
       character or punctuation mark),
     • the CFR's WO # is illegible, smudged, or cut off,
     • the CFR has no WO # printed on it.
   A CFR whose WO # doesn't match any preceding WO is an ORPHAN — it goes
   into no wo_document. The fact that it physically follows a WO is
   coincidental (staples come loose, scans get reordered) and does NOT
   justify bundling.
4. Plan / drawing pages (engineering site diagrams or plan sheets) belong to
   the Work Order they were scanned with. INCLUDE such a page in a wo_document
   when it falls within that WO's span in the stack — i.e. after that WO's WO
   page and before the next WO page (it usually sits after the WO's CFR, but
   include it wherever it appears within that span). These drawings must stay
   with the WO so the crew can reference them on site.
5. Never include blank or separator pages in any wo_document. Those pages get
   dropped.
6. Page numbers are 1-indexed.

Return ONLY this JSON shape (no explanation, no markdown fences):

{
  "wo_documents": [
    {
      "wo_id": "The WO # printed on the Work Order page (e.g. PT-11930 or RM-43281).",
      "pages": [list of 1-indexed page numbers for this WO's pages]
    }
  ]
}

If the PDF contains NO Work Order pages, return {"wo_documents": []}.`;

// ── Multi-page detection (Pass 1) ─────────────────────────────
async function detectWoDocuments(fileBytes: Buffer): Promise<{ wo_documents: { wo_id: string; pages: number[] }[] }> {
  const encoded = fileBytes.toString('base64');

  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': config.anthropic.apiKey,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify({
      model: 'claude-opus-4-6',
      max_tokens: 2048,
      messages: [{
        role: 'user',
        content: [
          { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: encoded } },
          { type: 'text', text: DETECT_WO_PROMPT },
        ],
      }],
    }),
  });

  if (!response.ok) {
    const err = await response.text();
    throw new Error(`Claude detect API error (${response.status}): ${err}`);
  }

  const result = await response.json() as any;
  let raw = result.content?.[0]?.text?.trim() || '';
  if (raw.startsWith('```')) {
    const lines = raw.split('\n');
    raw = lines.slice(1, lines[lines.length - 1] === '```' ? -1 : undefined).join('\n');
  }

  const parsed = JSON.parse(raw);
  if (!parsed?.wo_documents) throw new Error('Missing wo_documents key in detection response');

  // Normalize: dedupe pages, drop non-int, sort
  const normalized = [];
  for (const doc of (parsed.wo_documents || [])) {
    if (!doc || typeof doc !== 'object') continue;
    const seen = new Set<number>();
    const cleanPages: number[] = [];
    for (const p of (doc.pages || [])) {
      const pi = typeof p === 'number' ? p : parseInt(String(p), 10);
      if (isNaN(pi) || pi < 1 || seen.has(pi)) continue;
      seen.add(pi);
      cleanPages.push(pi);
    }
    if (cleanPages.length === 0) continue;
    normalized.push({
      wo_id: String(doc.wo_id || '').trim(),
      pages: cleanPages.sort((a, b) => a - b),
    });
  }

  return { wo_documents: normalized };
}

// ── PDF page count + splitting (using pdf-lib) ────────────────
async function pdfPageCount(fileBytes: Buffer): Promise<number> {
  try {
    const { PDFDocument } = await import('pdf-lib');
    const doc = await PDFDocument.load(fileBytes, { ignoreEncryption: true });
    return doc.getPageCount();
  } catch {
    return 0;
  }
}

async function splitPdfByPages(fileBytes: Buffer, pageLists: number[][]): Promise<Buffer[]> {
  const { PDFDocument } = await import('pdf-lib');
  const srcDoc = await PDFDocument.load(fileBytes, { ignoreEncryption: true });
  const totalPages = srcDoc.getPageCount();
  const outputs: Buffer[] = [];

  for (const pages of pageLists) {
    const newDoc = await PDFDocument.create();
    for (const p of pages) {
      const idx = p - 1; // 1-indexed → 0-indexed
      if (idx >= 0 && idx < totalPages) {
        const [copied] = await newDoc.copyPages(srcDoc, [idx]);
        newDoc.addPage(copied);
      }
    }
    const bytes = await newDoc.save();
    outputs.push(Buffer.from(bytes));
  }

  return outputs;
}

/**
 * Full scan pipeline — exact port of process_wo_scan from watch_and_fill.py.
 *
 * Page-count branching:
 *   ≤ 2 pages → single-WO parse (original behavior)
 *   > 2 pages → two-pass: detect WO boundaries, split PDF, parse each split
 */
export async function parseWorkOrderScanFull(fileBytes: Buffer, mimeType: string): Promise<ParsedWoData[]> {
  // Non-PDF → always single parse
  if (mimeType !== 'application/pdf') {
    return [await parseWorkOrderScan(fileBytes, mimeType)];
  }

  const pageCount = await pdfPageCount(fileBytes);

  // ≤ 2 pages → single WO
  if (pageCount <= 2) {
    console.log(`[Scan] Single-WO path (${pageCount} pages)`);
    return [await parseWorkOrderScan(fileBytes, mimeType)];
  }

  // > 2 pages → multi-WO pipeline
  console.log(`[Scan] Multi-WO path — Pass 1: detecting WO boundaries in ${pageCount} pages…`);

  let detectResult: { wo_documents: { wo_id: string; pages: number[] }[] };
  try {
    detectResult = await detectWoDocuments(fileBytes);
  } catch (err: any) {
    console.warn(`[Scan] Pass 1 detection failed: ${err.message}. Falling back to single-WO parse.`);
    return [await parseWorkOrderScan(fileBytes, mimeType)];
  }

  if (detectResult.wo_documents.length === 0) {
    console.warn('[Scan] Pass 1 returned zero WO documents — trying single-WO fallback.');
    return [await parseWorkOrderScan(fileBytes, mimeType)];
  }

  console.log(`[Scan] Pass 1: ${detectResult.wo_documents.length} WO document(s) detected`);
  for (const doc of detectResult.wo_documents) {
    console.log(`[Scan]   • ${doc.wo_id} on pages ${JSON.stringify(doc.pages)}`);
  }

  // Split PDF
  let splitList: Buffer[];
  try {
    splitList = await splitPdfByPages(fileBytes, detectResult.wo_documents.map(d => d.pages));
  } catch (err: any) {
    console.error(`[Scan] PDF split failed: ${err.message}. Falling back to single-WO parse.`);
    return [await parseWorkOrderScan(fileBytes, mimeType)];
  }

  // Pass 2: parse each split
  const results: ParsedWoData[] = [];
  for (let i = 0; i < splitList.length; i++) {
    const doc = detectResult.wo_documents[i];
    const hint = doc.wo_id || `part${i + 1}`;
    console.log(`[Scan] Pass 2: parsing split ${i + 1}/${splitList.length} (${hint})…`);
    try {
      const parsed = await parseWorkOrderScan(splitList[i], 'application/pdf');
      results.push(parsed);
    } catch (err: any) {
      console.error(`[Scan] Pass 2 parse failed for ${hint}: ${err.message}`);
      // Continue with remaining splits
    }
  }

  if (results.length === 0) {
    console.warn('[Scan] All Pass 2 parses failed — trying single-WO fallback on whole file.');
    return [await parseWorkOrderScan(fileBytes, mimeType)];
  }

  return results;
}

// ── Normalize (exact port of normalize_wo_data) ────────────────
function normalizeWoData(raw: any): ParsedWoData {
  // Borough
  const rawBorough = String(raw.borough || '').trim().toUpperCase();
  const regionCode = BOROUGH_MAP[rawBorough] || rawBorough;

  // Water blast logic
  let wbSqft = '';
  if (raw.water_blast_sqft != null) {
    const parsed = parseInt(String(raw.water_blast_sqft).replace(',', ''), 10);
    wbSqft = isNaN(parsed) ? String(raw.water_blast_sqft).trim() : String(parsed);
  }

  const remarks = String(raw.general_remarks || '').trim();
  const remarksLower = remarks.toLowerCase();
  const isMma = !!wbSqft || MMA_REMARK_KEYWORDS.some(kw => remarksLower.includes(kw));
  const hasRemarks = !!remarks;

  let waterBlastRequired: string;
  let waterBlastConfirmed: string;

  if (isMma) {
    waterBlastRequired = 'Yes - MMA';
    waterBlastConfirmed = wbSqft ? 'Yes' : 'No';
  } else if (hasRemarks) {
    waterBlastRequired = 'No - Thermo';
    waterBlastConfirmed = 'N/A';
    wbSqft = '';
  } else {
    waterBlastRequired = '';
    waterBlastConfirmed = 'N/A';
    wbSqft = '';
  }

  // Top markings
  const topMarkings: { category: string; description: string }[] = [];
  for (const m of (raw.top_markings || [])) {
    const category = String(m?.category || '').trim();
    const description = String(m?.description || '').trim();
    if (category && description) topMarkings.push({ category, description });
  }

  // Intersection grid
  const intersectionGrid: { intersection: string; n: string; e: string; s: string; w: string; stopMsg: string; stopLines: string }[] = [];
  for (const ig of (raw.intersection_grid || [])) {
    const intersection = String(ig?.intersection || '').trim();
    if (!intersection) continue;
    const cells = {
      n: String(ig?.n || '').trim(),
      e: String(ig?.e || '').trim(),
      s: String(ig?.s || '').trim(),
      w: String(ig?.w || '').trim(),
      stopMsg: String(ig?.stop_msg || '').trim(),
      stopLines: String(ig?.stop_lines || '').trim(),
    };
    if (Object.values(cells).some(v => v)) {
      intersectionGrid.push({ intersection, ...cells });
    }
  }

  // Bike lane markings
  const VALID_BIKE_TYPES = new Set(['Bike Symbol', 'Bike Arrow', 'Pedestrian Men']);
  let bikeLaneMarkings: { type: string; quantity: number | null; source: string }[] = [];
  for (const bm of (raw.bike_lane_markings || [])) {
    const btype = String(bm?.type || '').trim();
    if (!VALID_BIKE_TYPES.has(btype)) continue;
    let qty: number | null = null;
    if (bm?.quantity != null && String(bm.quantity).trim() !== '') {
      const parsed = parseInt(String(bm.quantity).trim(), 10);
      if (!isNaN(parsed)) qty = parsed;
    }
    bikeLaneMarkings.push({
      type: btype,
      quantity: qty,
      source: String(bm?.source || '').trim(),
    });
  }

  // Dedupe by type (prefer entry with quantity)
  const deduped = new Map<string, typeof bikeLaneMarkings[0]>();
  for (const bm of bikeLaneMarkings) {
    const existing = deduped.get(bm.type);
    if (!existing || (existing.quantity === null && bm.quantity !== null)) {
      deduped.set(bm.type, bm);
    }
  }
  bikeLaneMarkings = Array.from(deduped.values());

  // Work type derivation
  let workType: string;
  if (waterBlastRequired === 'Yes - MMA') {
    workType = 'MMA';
  } else if (intersectionGrid.length || topMarkings.length || bikeLaneMarkings.length) {
    workType = 'Thermo';
  } else {
    workType = '';
  }

  // WO-number prefix override
  const woPrefix = String(raw.work_order_id || '').trim().toUpperCase().split('-')[0];
  if (woPrefix === 'PT') {
    workType = 'MMA';
    waterBlastRequired = 'Yes - MMA';
    waterBlastConfirmed = wbSqft ? 'Yes' : 'No';
  } else if (woPrefix === 'RM') {
    workType = 'Thermo';
    waterBlastRequired = 'No - Thermo';
    waterBlastConfirmed = 'N/A';
    wbSqft = '';
  }

  return {
    workOrderId: String(raw.work_order_id || '').trim(),
    contractor: String(raw.contractor || '').trim().replace(/\b\w+/g, w => w[0].toUpperCase() + w.slice(1).toLowerCase()),
    contractNumber: String(raw.contract_number || '').trim(),
    regionCode,
    location: String(raw.location || '').trim().replace(/\b\w+/g, w => w[0].toUpperCase() + w.slice(1).toLowerCase()),
    fromStreet: String(raw.from_street || '').trim().replace(/\b\w+/g, w => w[0].toUpperCase() + w.slice(1).toLowerCase()),
    toStreet: String(raw.to_street || '').trim().replace(/\b\w+/g, w => w[0].toUpperCase() + w.slice(1).toLowerCase()),
    dueDate: String(raw.due_date || '').trim(),
    priority: String(raw.priority_level || '').trim(),
    pavementWorkType: String(raw.pavement_work_type || '').trim().toUpperCase(),
    woReceivedDate: String(raw.wo_received_date || '').trim(),
    dateEntered: String(raw.date_entered || '').trim(),
    school: String(raw.school || 'NA').trim() || 'NA',
    prepBy: String(raw.prep_by || '').trim(),
    waterBlastRequired,
    waterBlastConfirmed,
    waterBlastSqft: String(wbSqft),
    workType,
    generalRemarks: remarks,
    topMarkings,
    intersectionGrid,
    bikeLaneMarkings,
  };
}
