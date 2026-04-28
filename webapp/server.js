/**
 * Oneiro Ops — Express Backend
 *
 * In production (Railway): serves the Vite-built React app from ./dist
 *   and proxies all /api/* calls to the Apps Script doPost endpoint.
 *
 * In development: only handles /api routes; Vite dev server handles the frontend
 *   and proxies /api → this server via vite.config.js proxy setting.
 *
 * Environment variables (set in Railway):
 *   PORT                     — Railway sets this automatically
 *   APPS_SCRIPT_URL          — the deployed Apps Script web app URL
 *   APPS_SCRIPT_KEY          — the UPLOAD_SECRET script property value
 *   NODE_ENV                 — "production" | "development"
 */

import express  from 'express'
import cors     from 'cors'
import path     from 'path'
import multer   from 'multer'
import { fileURLToPath } from 'url'
import 'dotenv/config'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const app  = express()
const PORT = process.env.PORT || 3001

// ── Middleware ────────────────────────────────────────────────
app.use(cors())
app.use(express.json({ limit: '5mb' }))

// multer: stores uploads in memory, 15MB per file, 20 files max per request
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024, files: 20 }
})


// ── Apps Script proxy helper ──────────────────────────────────
async function callAppsScript(action, data = null) {
  const url = process.env.APPS_SCRIPT_URL
  const key = process.env.APPS_SCRIPT_KEY

  if (!url || !key) {
    throw new Error('APPS_SCRIPT_URL or APPS_SCRIPT_KEY env var not set')
  }

  const body = { action, key, ...(data ? { data } : {}) }

  const res = await fetch(url, {
    method:  'POST',
    headers: { 'Content-Type': 'application/json' },
    body:    JSON.stringify(body)
  })

  // Apps Script normally returns 200 + JSON. If the deployment is in a
  // bad state (fresh deploy that needs re-auth, URL pointing at a stale
  // version, Google outage), it returns an HTML login/error page and
  // res.json() throws a cryptic "Unexpected token '<'" that bubbles to
  // the browser as a confusing JSON-parse error. Catch that shape and
  // raise something the user can actually act on.
  const text = await res.text()
  const trimmed = text.trim()
  if (trimmed.startsWith('<')) {
    const status = res.status
    const redirected = res.redirected ? ' (redirected — Apps Script may need re-authorization)' : ''
    throw new Error(
      `Apps Script returned HTML instead of JSON${redirected} — ` +
      `action="${action}", status=${status}. ` +
      `This usually means the Web App deployment needs to be re-authorized ` +
      `or APPS_SCRIPT_URL is pointing at a stale deployment.`
    )
  }
  let json
  try {
    json = JSON.parse(text)
  } catch (e) {
    throw new Error(
      `Apps Script response wasn't JSON (action="${action}", status=${res.status}): ` +
      `${text.slice(0, 200)}`
    )
  }
  if (json.error) throw new Error(json.error)
  return json
}


// ── API Routes ────────────────────────────────────────────────

/**
 * GET /api/health
 * Simple health check — Railway uses this to verify the service is up.
 */
app.get('/api/health', (_req, res) => {
  res.json({ ok: true, service: 'oneiro-ops-webapp', ts: new Date().toISOString() })
})

/**
 * GET /api/wos
 * Returns all active (non-complete) Work Orders for the field report dropdown.
 */
app.get('/api/wos', async (_req, res) => {
  try {
    const data = await callAppsScript('get_active_wos')
    res.json(data)
  } catch (err) {
    console.error('GET /api/wos error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/dashboard
 * Returns full WO list + stats for the admin dashboard.
 */
app.get('/api/dashboard', async (_req, res) => {
  try {
    const data = await callAppsScript('get_dashboard_data')
    res.json(data)
  } catch (err) {
    console.error('GET /api/dashboard error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/wo-markings/:woId
 * Returns the Marking Items rows pre-populated for this WO (from the
 * scan) plus any completions that have already been written. The Field
 * Report UI uses this to render the list of items the crew needs to
 * measure on-site.
 */
app.get('/api/wo-markings/:woId', async (req, res) => {
  try {
    const data = await callAppsScript('get_marking_items', { wo_id: req.params.woId })
    res.json(data)
  } catch (err) {
    console.error('GET /api/wo-markings error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/marking-items
 * Create ONE manually-added marking item. Body: full item fields (wo_id,
 * category, etc.). Response: { item: <full row object> }.
 */
app.post('/api/marking-items', async (req, res) => {
  try {
    const data = await callAppsScript('create_marking_item', req.body)
    res.json(data)
  } catch (err) {
    console.error('POST /api/marking-items error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * PATCH /api/marking-items/:itemId
 * Update any subset of editable fields on one marking item. Body: a
 * partial patch (quantity, unit, category, etc.). Response:
 * { item: <full updated row object> }.
 */
app.patch('/api/marking-items/:itemId', async (req, res) => {
  try {
    const data = await callAppsScript('update_marking_item', {
      ...req.body,
      item_id: req.params.itemId,
    })
    res.json(data)
  } catch (err) {
    console.error('PATCH /api/marking-items error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * DELETE /api/marking-items
 * Delete one or more marking items by Item ID. Body: { item_ids: [...] }.
 * Response: { deleted: [<ids removed>] }.
 */
app.delete('/api/marking-items', async (req, res) => {
  try {
    const data = await callAppsScript('delete_marking_items', req.body)
    res.json(data)
  } catch (err) {
    console.error('DELETE /api/marking-items error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/field-report
 * Submits a crew field report (writes to Daily Sign-In Data + WO Tracker).
 * Body: the field report data object (see Apps Script handleSubmitFieldReport_)
 */
app.post('/api/field-report', async (req, res) => {
  try {
    const data = await callAppsScript('submit_field_report', req.body)
    res.json(data)
  } catch (err) {
    console.error('POST /api/field-report error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/tools/daily-documents
 * Runs the daily-document generator. Body: { date?: 'YYYY-MM-DD' }.
 * Empty/missing date → today.  Replaces the "Generate Daily Documents"
 * items on the spreadsheet's custom menu, which stopped firing
 * reliably from standalone-script installable onOpen.
 */
app.post('/api/tools/daily-documents', async (req, res) => {
  try {
    const { date } = req.body || {}
    // Convert ISO YYYY-MM-DD → MM/DD/YYYY so Apps Script's new Date()
    // parses it as LOCAL midnight. ISO strings get parsed as UTC per
    // the spec, which shifted "today" back by one day in Eastern Time
    // and caused "no sign-in entries" errors.
    let normalised = ''
    if (date && /^\d{4}-\d{2}-\d{2}$/.test(String(date))) {
      const [y, m, d] = String(date).split('-')
      normalised = `${m}/${d}/${y}`
    } else if (date) {
      normalised = String(date)
    }
    const data = await callAppsScript('generate_daily_documents', { date: normalised })
    res.json(data)
  } catch (err) {
    console.error('POST /api/tools/daily-documents error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/tools/certified-payroll
 * Runs the certified-payroll generator for a given Monday-starting
 * week. Body: { week_start: 'YYYY-MM-DD' }.  Converted to MM/DD/YYYY
 * here so the Apps Script handler can parse it verbatim.
 */
app.post('/api/tools/certified-payroll', async (req, res) => {
  try {
    const { week_start } = req.body || {}
    if (!week_start || !/^\d{4}-\d{2}-\d{2}$/.test(String(week_start))) {
      return res.status(400).json({ error: 'Missing or malformed week_start (expected YYYY-MM-DD)' })
    }
    const [y, m, d] = String(week_start).split('-')
    const mmddyyyy = `${m}/${d}/${y}`
    const data = await callAppsScript('generate_certified_payroll', { week_start: mmddyyyy })
    res.json(data)
  } catch (err) {
    console.error('POST /api/tools/certified-payroll error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/tools/process-approved-docs
 * Manually fire the same function the 10-min cron runs — emails any
 * pending doc in Approved Docs + archives it (or moves to Archive
 * Errors on failure). Script-wide lock inside Apps Script means it's
 * safe to fire concurrently with the cron.
 */
app.post('/api/tools/process-approved-docs', async (_req, res) => {
  try {
    const data = await callAppsScript('process_approved_documents')
    res.json(data)
  } catch (err) {
    console.error('POST /api/tools/process-approved-docs error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/waterblasting/:woId/confirm
 * Flips the "Water Blast Confirmed?" flag on the Work Order Tracker
 * (col N). MMA jobs can't have a Field Report submitted until this is
 * "Yes" — enforced both client-side (greyed UI) and server-side (guard
 * inside handleSubmitFieldReport_).
 *
 * Body: { confirmed: boolean }
 */
app.post('/api/waterblasting/:woId/confirm', async (req, res) => {
  try {
    const data = await callAppsScript('set_waterblast_confirmed', {
      wo_id:     req.params.woId,
      confirmed: !!req.body?.confirmed,
    })
    res.json(data)
  } catch (err) {
    console.error('POST /api/waterblasting/:woId/confirm error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/field-report/finalize
 * Triggers Sign-In + CFR JSON generation AFTER the submit returned
 * success. The client fires this as fire-and-forget so the user sees
 * the success screen immediately instead of waiting for Drive writes.
 */
app.post('/api/field-report/finalize', async (req, res) => {
  try {
    const data = await callAppsScript('finalize_field_report_docs', req.body)
    res.json(data)
  } catch (err) {
    console.error('POST /api/field-report/finalize error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/upload-wo
 * Uploads one scanned WO file (PDF/JPEG/PNG) to Drive's Scan Inbox
 * folder so the Railway watcher picks it up and runs the existing
 * Claude Vision parse + write_wo pipeline.
 * Multipart form field:
 *   file  — the WO scan (PDF or image)
 */
app.post('/api/upload-wo', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file attached' })

    const base64 = req.file.buffer.toString('base64')
    const data   = await callAppsScript('upload_wo_scan', {
      filename:  req.file.originalname,
      mime_type: req.file.mimetype,
      data:      base64
    })
    res.json(data)
  } catch (err) {
    console.error('POST /api/upload-wo error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/scan-status
 * Batched per-file status lookup for the Scan WO page's polling loop.
 * Body: { file_ids: [<string>, ...] }
 * Response: { statuses: [{ file_id, status, wo_ids?, message? }, ...] }
 */
app.post('/api/scan-status', async (req, res) => {
  try {
    const fileIds = Array.isArray(req.body?.file_ids) ? req.body.file_ids : []
    if (fileIds.length === 0) return res.json({ statuses: [] })
    const data = await callAppsScript('get_scan_status', { file_ids: fileIds })
    res.json(data)
  } catch (err) {
    console.error('POST /api/scan-status error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/approvals
 * Returns every PDF in Docs Needing Review subfolders, grouped per
 * doc type with subtitle + created_at metadata. Webapp Approvals page
 * uses this as its master list.
 */
app.get('/api/approvals', async (_req, res) => {
  try {
    const data = await callAppsScript('list_pending_approvals')
    res.json(data)
  } catch (err) {
    console.error('GET /api/approvals error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/approvals/:fileId/pdf
 * Streams the raw PDF bytes for a pending-approval file. Apps Script
 * returns base64; we decode and send as application/pdf so react-pdf
 * can consume the URL directly.
 */
app.get('/api/approvals/:fileId/pdf', async (req, res) => {
  try {
    const data = await callAppsScript('get_drive_file_bytes',
                                      { file_id: req.params.fileId })
    if (!data || !data.data) {
      return res.status(500).json({ error: 'Apps Script returned no bytes' })
    }
    const buf = Buffer.from(data.data, 'base64')
    res.setHeader('Content-Type', data.mime_type || 'application/pdf')
    res.setHeader('Content-Length', buf.length)
    res.setHeader('Cache-Control', 'no-store')
    res.send(buf)
  } catch (err) {
    console.error(`GET /api/approvals/${req.params.fileId}/pdf error:`, err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/approvals/:fileId/approve
 * Moves the file from Docs Needing Review → ✅ Approved Docs.
 * The existing processApprovedDocuments cron picks it up on its next
 * tick (≤ 10 min) and handles email + archive identically to a manual
 * Drive drag.
 */
app.post('/api/approvals/:fileId/approve', async (req, res) => {
  try {
    const data = await callAppsScript('approve_doc',
                                      { file_id: req.params.fileId })
    res.json(data)
  } catch (err) {
    console.error(`POST /api/approvals/${req.params.fileId}/approve error:`, err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/approvals/:fileId/skip-signoff
 * Approves a document WITHOUT the pdf-lib principal-signature overlay.
 * Used for manually-uploaded sign-in sheets (filename ends in
 * `_MANUAL.pdf`) that the principal already wet-signed by hand — we
 * don't want to drop a second electronic signature on top.
 */
app.post('/api/approvals/:fileId/skip-signoff', async (req, res) => {
  try {
    const data = await callAppsScript('approve_doc_skip_signoff',
                                      { file_id: req.params.fileId })
    res.json(data)
  } catch (err) {
    console.error(`POST /api/approvals/${req.params.fileId}/skip-signoff error:`, err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/approvals/:fileId/approve-signin
 * Sign-In-specific approval: takes the principal's signature + printed
 * name + title from the webapp modal, patches the PDF via pdf-lib
 * (text fields + signature image overlay), and hands the patched
 * bytes to Apps Script for upload to Approved Docs + trash the unsigned
 * original. The cron handles email + archive downstream.
 *
 * Body: { signature_b64: string, name: string, title: string }
 */
app.post('/api/approvals/:fileId/approve-signin', express.json({ limit: '5mb' }), async (req, res) => {
  const fileId = req.params.fileId
  try {
    const { signature_b64, name, title } = req.body || {}
    if (!signature_b64) return res.status(400).json({ error: 'Missing signature_b64' })
    if (!name || !String(name).trim())   return res.status(400).json({ error: 'Missing name' })
    if (!title || !String(title).trim()) return res.status(400).json({ error: 'Missing title' })

    // 1. Fetch original PDF bytes + filename from Drive
    const fetched = await callAppsScript('get_drive_file_bytes', { file_id: fileId })
    if (!fetched || !fetched.data) {
      return res.status(500).json({ error: 'Couldn\u2019t fetch PDF bytes from Drive' })
    }
    const pdfBytes  = Buffer.from(fetched.data, 'base64')
    const filename  = fetched.filename

    // 2. Patch with pdf-lib
    const { PDFDocument } = await import('pdf-lib')
    const pdfDoc = await PDFDocument.load(pdfBytes, { updateMetadata: false })
    const form   = pdfDoc.getForm()

    // Date: today, M/D/YY in the spreadsheet's timezone (America/New_York).
    // Without the TZ override Node runs in UTC on Railway, which rolls over
    // to "tomorrow" at 20:00 ET and causes the printed date to be wrong.
    const parts = new Intl.DateTimeFormat('en-US', {
      timeZone:  'America/New_York',
      month:     'numeric',
      day:       'numeric',
      year:      '2-digit',
    }).formatToParts(new Date())
    const pick = (t) => parts.find(p => p.type === t)?.value || ''
    const dateStr = `${pick('month')}/${pick('day')}/${pick('year')}`

    // Fill text fields
    const trySetText = (fieldName, value) => {
      try { form.getTextField(fieldName).setText(String(value || '')) }
      catch (e) { console.warn(`pdf-lib setText failed on ${fieldName}: ${e.message}`) }
    }
    trySetText('Contractor_Name',      name)
    trySetText('Contractor_Title',     title)
    trySetText('Date_Signature_Block', dateStr)

    // Overlay the signature PNG onto the Contractor_Signature field's rect.
    // Matches workers/fill_signin.py's crew-leader overlay: expand the
    // rect by (h=1, v=3) points, then scale the signature to fit the
    // expanded box while preserving aspect ratio, centered. The client
    // already cropped the PNG to its ink bounding box, so the visible
    // ink fills the space well.
    const LEADER_SIG_H_PAD = 1.0
    const LEADER_SIG_V_PAD = 3.0
    const pngB64 = String(signature_b64).replace(/^data:image\/png;base64,/, '')
    const pngBytes = Buffer.from(pngB64, 'base64')
    const sigImage = await pdfDoc.embedPng(pngBytes)
    try {
      const sigField  = form.getField('Contractor_Signature')
      const widgets   = sigField.acroField.getWidgets()
      const widget    = widgets[0]
      const rect      = widget.getRectangle()
      const pageRef = widget.P()
      const pages = pdfDoc.getPages()
      const page = pages.find(p => p.ref === pageRef) || pages[0]

      const boxX = rect.x      - LEADER_SIG_H_PAD
      const boxY = rect.y      - LEADER_SIG_V_PAD
      const boxW = rect.width  + 2 * LEADER_SIG_H_PAD
      const boxH = rect.height + 2 * LEADER_SIG_V_PAD

      const scale = Math.min(boxW / sigImage.width, boxH / sigImage.height)
      const drawW = sigImage.width  * scale
      const drawH = sigImage.height * scale
      const drawX = boxX + (boxW - drawW) / 2
      const drawY = boxY + (boxH - drawH) / 2
      page.drawImage(sigImage, { x: drawX, y: drawY, width: drawW, height: drawH })
    } catch (e) {
      console.warn('pdf-lib signature overlay failed:', e.message)
      // Fall through — name/title/date are still applied
    }

    // Flatten the form so name/title/date render reliably everywhere
    // (pdf-lib writes /V, but we want no interactive fields remaining
    // after principal signature either).
    try { form.flatten() }
    catch (e) { console.warn('pdf-lib form.flatten failed:', e.message) }

    const signedBytes = await pdfDoc.save({ updateFieldAppearances: false })

    // 3. Upload patched bytes to Approved Docs + trash original via Apps Script
    const result = await callAppsScript('approve_signin_with_bytes', {
      file_id:   fileId,
      filename:  filename,
      bytes_b64: Buffer.from(signedBytes).toString('base64'),
    })
    if (result.error) throw new Error(result.error)
    res.json(result)
  } catch (err) {
    console.error(`POST /api/approvals/${fileId}/approve-signin error:`, err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/signin-queue/parse-upload
 * Accepts a single PDF (≤ 15 MB, multipart "file") that a crew leader
 * filled by hand and uploaded for one queue entry. Calls Claude Vision
 * with a strict-JSON prompt and returns the parsed crew rows so the
 * Sign-In tab can pre-fill its form for the user to confirm.
 *
 * Required env: ANTHROPIC_API_KEY (the same key the Python worker uses).
 */
app.post('/api/signin-queue/parse-upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file)            return res.status(400).json({ error: 'No file attached' })
    if (!process.env.ANTHROPIC_API_KEY) {
      return res.status(500).json({ error: 'ANTHROPIC_API_KEY not configured on this server' })
    }
    const mime = req.file.mimetype || 'application/pdf'
    if (mime !== 'application/pdf') {
      return res.status(400).json({ error: 'Only PDF uploads are supported (got ' + mime + ')' })
    }

    // Orientation handling. Two pdf-lib passes:
    //   1. Strip any existing /Rotate metadata so we work from a known
    //      baseline (raw drawn content).
    //   2. If the raw page is portrait (height > width), the paper was
    //      almost certainly a landscape sheet scanned sideways — set
    //      /Rotate=270 (= 90° CCW display rotation, which un-rotates
    //      the 90° CW content drawn within a portrait page). Setting
    //      /Rotate=90 would add ANOTHER 90° CW rotation and end up
    //      upside-down, which is the bug we hit before.
    //
    // We do this BEFORE sending to Claude so the OCR pass sees an
    // upright page, which keeps row order natural and removes any
    // prompt complexity around orientation.
    let workingBytes = req.file.buffer
    let rotatedAny   = false
    try {
      const { PDFDocument, degrees } = await import('pdf-lib')
      const pdfDoc = await PDFDocument.load(req.file.buffer, { updateMetadata: false })
      pdfDoc.getPages().forEach(page => {
        page.setRotation(degrees(0))
        const { width, height } = page.getSize()
        if (height > width) {
          page.setRotation(degrees(270))
          rotatedAny = true
        }
      })
      workingBytes = Buffer.from(await pdfDoc.save())
    } catch (e) {
      console.warn('Sign-in upload: rotation pre-pass failed:', e.message)
      workingBytes = req.file.buffer
    }
    const base64 = workingBytes.toString('base64')

    // Pull the Employee Registry from Apps Script so we can hand
    // Claude the closed list of valid names. The prompt then asks
    // Claude to return one of THOSE values exactly — handwriting
    // matching happens inside the model rather than in our code, so
    // unusual spelling/cursive variants still resolve correctly.
    let employeeList = []
    try {
      const empRes = await callAppsScript('list_employees')
      employeeList = (empRes && Array.isArray(empRes.employees))
        ? empRes.employees.map(e => String(e.name || '').trim()).filter(Boolean)
        : []
    } catch (e) {
      console.warn('Sign-in upload: could not fetch employee list:', e.message)
    }
    const employeeListText = employeeList.length
      ? employeeList.map(n => `  - ${n}`).join('\n')
      : '  (registry currently empty — return the raw handwriting)'

    const prompt = `You are extracting information from a hand-filled construction crew Sign-In Sheet (NYC DOT "EMPLOYEES' DAILY SIGN-IN LOG"). The sheet is filled in by hand by a crew leader at the start and end of a shift. Return ONLY valid JSON — no markdown fences, no commentary.

Schema:
{
  "crew": [
    {
      "name":           "<EXACTLY one name from the Employee Registry below — see Name matching rules.>",
      "classification": "<Either \\"LP\\" or \\"SAT\\". LP = Line Person/Crew Chief. SAT = Stripe Assistant Technician. Return null if blank.>",
      "time_in":        "<24-hour HH:MM from the \\"Time In\\" column. See \\"Time interpretation\\" below. null if unreadable.>",
      "time_out":       "<24-hour HH:MM from the \\"Time Out\\" column. See \\"Time interpretation\\" below. null if unreadable.>"
    }
  ],
  "contractor_name":  "<Printed name at bottom of sheet (Contractor's Representative), or null>",
  "contractor_title": "<Title at bottom of sheet, or null>",
  "date_inferred":    "<Date written at the top of the sheet in YYYY-MM-DD if parseable, otherwise null>"
}

Employee Registry (the ONLY valid values for crew[].name):
${employeeListText}

Name matching rules:
- For each handwritten name on the sheet, find the SINGLE registry entry above whose spelling is closest to the handwriting (account for cursive, partial first/last names, smudges, abbreviations, swapped first/last name order, common misspellings — pick the best fit even if not a perfect match).
- Return that registry entry verbatim — exact capitalization and spelling as listed above.
- The crew is drawn from this list, so a real crew row WILL match. Do not return null, do not invent a name not in the list.
- Only return null for the name if the row is completely blank (no handwriting at all).

Time interpretation (CRITICAL — most common error):
- This is night shift work for road striping crews. Time In is typically late evening (8pm–11pm), Time Out is typically early morning (4am–7am).
- If a row shows Time In near 10–11 with a "P" or "PM" suffix, output "22:00" or "23:00" — never the AM equivalent.
- If a row shows Time Out near 5–7 with an "A" or "AM" suffix, output "05:00"–"07:00".
- If the AM/PM marker is illegible BUT the value is consistent with a typical night-shift pattern (e.g. Time In ~22:00, Time Out ~06:00), favor that interpretation over a daytime one.
- "11PM" → "23:00"; "11AM" → "11:00". Don't conflate them.
- Convert 12-hour formats: "10pm" → "22:00", "6 AM" → "06:00", "10:30PM" → "22:30".

Column layout (CRITICAL — must not be confused):
- The form has these columns left-to-right: Employees Name | Classification | Time In | Employees Signature | Time Out | Employees Signature.
- Time In is to the LEFT of the signature column; Time Out is to the RIGHT of that signature column.
- For each row, return the time written in the "Time In" column as time_in and the time written in the "Time Out" column as time_out. Do NOT swap them.

Rules:
- Skip rows that are entirely blank. Only include crew rows that have at least one filled cell (name, time, or classification).
- Classification: only "LP" or "SAT". If the sheet uses other codes, pick the closest match; otherwise null.
- Never wrap the JSON in markdown fences.`

    const apiRes = await fetch('https://api.anthropic.com/v1/messages', {
      method:  'POST',
      headers: {
        'x-api-key':         process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'content-type':      'application/json',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5',
        max_tokens: 2048,
        messages: [{
          role: 'user',
          content: [
            {
              type: 'document',
              source: {
                type:       'base64',
                media_type: 'application/pdf',
                data:       base64,
              },
            },
            { type: 'text', text: prompt },
          ],
        }],
      }),
    })

    if (!apiRes.ok) {
      const errText = await apiRes.text()
      throw new Error(`Anthropic API error ${apiRes.status}: ${errText.slice(0, 300)}`)
    }
    const apiJson = await apiRes.json()
    const textBlock = (apiJson.content || []).find(b => b.type === 'text')
    if (!textBlock) throw new Error('Anthropic response had no text block')

    // Be defensive: Claude is told no markdown fences but occasionally
    // wraps JSON in ```json ... ```. Strip if present.
    let raw = String(textBlock.text || '').trim()
    raw = raw.replace(/^```(?:json)?\s*/i, '').replace(/```$/, '').trim()
    let parsed
    try {
      parsed = JSON.parse(raw)
    } catch (e) {
      throw new Error('Claude returned non-JSON text — first 200 chars: ' + raw.slice(0, 200))
    }

    // Echo the (already-rotated) bytes back to the client so they go
    // straight into /api/signin without re-uploading. The archived PDF
    // is then the user-readable landscape version, not the original
    // sideways scan.
    res.json({
      parsed,
      rotated: rotatedAny,
      upload: {
        filename:    req.file.originalname,
        mime_type:   mime,
        size:        workingBytes.length,
        data_b64:    base64,
      },
    })
  } catch (err) {
    console.error('POST /api/signin-queue/parse-upload error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/employees
 * Returns the list of employee names (Employee Registry col B) for the
 * Sign-In tab's crew dropdown. Apps Script side caches for 5 min.
 */
app.get('/api/employees', async (_req, res) => {
  try {
    const data = await callAppsScript('list_employees')
    res.json(data)
  } catch (err) {
    console.error('GET /api/employees error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/signin-queue
 * Returns outstanding (date, contract) groups derived from the Work Day
 * Log. Each group lists the WOs that need a sign-in for that contract on
 * that date.
 */
app.get('/api/signin-queue', async (_req, res) => {
  try {
    const data = await callAppsScript('list_signin_queue')
    res.json(data)
  } catch (err) {
    console.error('GET /api/signin-queue error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/signin-queue/check-continuation
 * Sanity check before a Sign-In submit. Given the user's chosen
 * Time-In + default date, asks Apps Script whether there's a
 * recent Daily Sign-In Data row whose Time-Out fell within 60 minutes
 * before this submission. The Sign-In tab uses the response to decide
 * whether to prompt the user "is this a continuation of last night's
 * shift?" before posting the actual sign-in.
 * Body: { time_in: "HH:MM", default_date: "YYYY-MM-DD" }
 */
app.post('/api/signin-queue/check-continuation', async (req, res) => {
  try {
    const data = await callAppsScript('check_signin_continuation', req.body)
    res.json(data)
  } catch (err) {
    console.error('POST /api/signin-queue/check-continuation error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/signin-queue/day-hours
 * Returns a map of employee → total hours already on Daily Sign-In
 * Data for a given operational day. The Sign-In tab uses this to
 * compose "Shift Totals" — i.e. what the worker's day will look like
 * once the in-flight sign-in is added on top.
 * Body: { date: "YYYY-MM-DD" }
 * Response: { totals: { "<name>": <hours>, ... } }
 */
app.post('/api/signin-queue/day-hours', async (req, res) => {
  try {
    const data = await callAppsScript('list_signin_day_hours', req.body)
    res.json(data)
  } catch (err) {
    console.error('POST /api/signin-queue/day-hours error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/signin
 * Submits a Sign-In sheet for one queue entry. Body:
 *   { queue_id, date, contract_number, borough, contractor,
 *     wo_ids: [...], crew: [...],
 *     contractor_name, contractor_title, date_signed, contractor_signature_b64,
 *     source: 'generated' | 'uploaded',
 *     upload_blob_b64?, upload_filename? }
 * The Apps Script side appends Daily Sign-In Data rows, marks the Work
 * Day Log entries Submitted, and either drops a JSON for the Python
 * filler ('generated') or stores the uploaded PDF directly ('uploaded').
 */
app.post('/api/signin', async (req, res) => {
  try {
    const data = await callAppsScript('submit_signin', req.body)
    res.json(data)
  } catch (err) {
    console.error('POST /api/signin error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/scan-uploads-today
 * Returns today's scan-originated WO Tracker rows grouped by upload.
 * The Scan WO page uses this as its source of truth so the queue
 * reflects the tracker (cross-device, survives browser clears,
 * respects admin deletions).
 */
app.get('/api/scan-uploads-today', async (_req, res) => {
  try {
    const data = await callAppsScript('get_scan_uploads_today')
    res.json(data)
  } catch (err) {
    console.error('GET /api/scan-uploads-today error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/upload-photo
 * Uploads one site photo to Drive inside the WO's Photos folder.
 * Multipart form fields:
 *   wo_id  — Work Order # (text)
 *   photo  — image file
 */
app.post('/api/upload-photo', upload.single('photo'), async (req, res) => {
  try {
    if (!req.file)        return res.status(400).json({ error: 'No file attached' })
    if (!req.body.wo_id)  return res.status(400).json({ error: 'wo_id is required' })

    const base64 = req.file.buffer.toString('base64')
    const data   = await callAppsScript('upload_photo', {
      wo_id:     req.body.wo_id,
      filename:  req.file.originalname,
      mime_type: req.file.mimetype,
      data:      base64
    })
    res.json(data)
  } catch (err) {
    console.error('POST /api/upload-photo error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/upload-signature
 * Uploads a crew member's digital signature PNG to Drive.
 * Body (JSON):
 *   wo_id      — Work Order #
 *   crew_name  — employee full name
 *   signature  — "time_in" | "time_out"
 *   work_date  — YYYY-MM-DD
 *   data       — base64-encoded PNG
 */
app.post('/api/upload-signature', async (req, res) => {
  try {
    const data = await callAppsScript('upload_signature', req.body)
    res.json(data)
  } catch (err) {
    console.error('POST /api/upload-signature error:', err.message)
    res.status(500).json({ error: err.message })
  }
})


// ── Static file serving (production only) ────────────────────
if (process.env.NODE_ENV === 'production') {
  const distDir = path.join(__dirname, 'dist')
  app.use(express.static(distDir))

  // All non-API routes → serve index.html (client-side routing)
  app.get('*', (req, res) => {
    if (!req.path.startsWith('/api')) {
      res.sendFile(path.join(distDir, 'index.html'))
    }
  })
}


// ── Start ─────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`🚀 Oneiro Ops web server running on port ${PORT}`)
  console.log(`   Mode: ${process.env.NODE_ENV || 'development'}`)
  if (process.env.NODE_ENV !== 'production') {
    console.log(`   API:  http://localhost:${PORT}/api/health`)
    console.log(`   App:  http://localhost:5173  (Vite dev server)`)
  }
})
