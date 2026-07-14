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
import archiver from 'archiver'
import crypto   from 'crypto'
import { fileURLToPath } from 'url'
import { monitorEventLoopDelay } from 'node:perf_hooks'
import 'dotenv/config'

import {
  qbConfigStatus, buildAuthorizeUrl, exchangeAuthCode,
  findCustomerByName, createInvoiceForWO, checkQbConnection,
  buildInvoiceViewUrl, qbInvoiceExists,
} from './server/qb.js'
import { assertQbItemsConfigured } from './server/qbItems.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const app  = express()
const PORT = process.env.PORT || 3001

// ── Middleware ────────────────────────────────────────────────
app.use(cors())
app.use(express.json({ limit: '5mb' }))

// Disable HTTP TRACE explicitly — Intuit security requirement.
// Express defaults to 404 on unrecognized methods, but an explicit
// 405 makes the intent visible and traceable in logs.
app.use((req, res, next) => {
  if (req.method === 'TRACE') return res.status(405).send('Method Not Allowed')
  next()
})

// Tighten cache-control on every /api/* response — Intuit security
// requirement: SSL pages and pages with sensitive data must use
// `no-cache, no-store` (not `private`) in the Cache-Control header.
// Static files (legal pages, Vite-built React bundles with content-
// hashed filenames) keep their default caching since they're public.
app.use((req, res, next) => {
  if (req.path.startsWith('/api/')) {
    res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private')
    res.set('Pragma', 'no-cache')
    res.set('Expires', '0')
  }
  next()
})

// multer: stores uploads in memory, 20MB per file, 20 files max per request
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024, files: 20 }
})


// ── Event-loop lag monitor ────────────────────────────────────
// callAppsScript measures its own latency with Date.now() around an
// await, which runs on the event loop. If this process is CPU-blocked
// (pdf-lib parses synchronously; zlib eats the libuv threadpool that
// DNS also uses), that timer inflates even when Apps Script answered
// promptly — so `elapsed` alone cannot tell "Google was slow" apart
// from "we couldn't get back to the socket". This histogram can.
//
// Reading the two together:
//   high elapsed + high elMax  → this Node process was blocked
//   high elapsed + low  elMax  → Apps Script really was slow
const loopLag = monitorEventLoopDelay({ resolution: 10 })
loopLag.enable()

// In-flight counters, logged with each call so a slow request can be
// correlated with whatever else was running at the time.
let asInFlight = 0
let batchDownloadsActive = 0

const ms = (ns) => Math.round(ns / 1e6)


// ── Apps Script proxy helper ──────────────────────────────────
async function callAppsScript(action, data = null) {
  const url = process.env.APPS_SCRIPT_URL
  const key = process.env.APPS_SCRIPT_KEY

  if (!url || !key) {
    throw new Error('APPS_SCRIPT_URL or APPS_SCRIPT_KEY env var not set')
  }

  const body    = { action, key, ...(data ? { data } : {}) }
  const reqJson = JSON.stringify(body)

  // ── Diagnostic instrumentation ──────────────────────────────
  // Characterises the "HTML instead of JSON" failures. Per call we log:
  // how long the round trip took, split into time-to-headers (ttfb) vs
  // time spent draining the body; peak event-loop lag during the call;
  // how many other Apps Script calls and zip streams were in flight; the
  // final URL after redirects; content-type + size; and, when HTML comes
  // back, the human-readable Google error text. All lines are prefixed
  // [AS] for easy grep in the Railway logs.
  //
  // NOTE: no timeout and no retry here, deliberately. A blanket 28s
  // AbortSignal was tried and reverted (644d17a) — list_documents_for_batch
  // legitimately runs longer, so the timeout turned a slow-but-working
  // call into a guaranteed failure.
  const started = Date.now()
  loopLag.reset()
  asInFlight++
  let res, text, tHeaders
  try {
    res = await fetch(url, {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    reqJson
    })
    tHeaders = Date.now()
    text = await res.text()
  } finally {
    asInFlight--
  }

  const finished = Date.now()
  const elapsed  = finished - started
  const ttfbMs   = tHeaders - started
  const bodyMs   = finished - tHeaders
  const elMaxMs  = ms(loopLag.max)
  const elMeanMs = Number.isFinite(loopLag.mean) ? ms(loopLag.mean) : 0
  const ctype    = res.headers.get('content-type') || ''
  const clen     = res.headers.get('content-length') || ''
  const trimmed  = text.trim()
  const looksHtml = trimmed.startsWith('<')

  console.log(
    `[AS] action=${action} status=${res.status} redirected=${res.redirected} ` +
    `elapsed=${elapsed}ms ttfb=${ttfbMs}ms body=${bodyMs}ms ` +
    `elMax=${elMaxMs}ms elMean=${elMeanMs}ms ` +
    `asInFlight=${asInFlight} zips=${batchDownloadsActive} ` +
    `reqBytes=${reqJson.length} respBytes=${text.length} ` +
    `contentLength=${clen || 'n/a'} contentType="${ctype}" ` +
    `finalUrl=${res.url} kind=${looksHtml ? 'HTML' : 'JSON/other'}`
  )

  // Apps Script normally returns 200 + JSON. If the deployment is in a
  // bad state (fresh deploy that needs re-auth, URL pointing at a stale
  // version, Google outage), it returns an HTML login/error page and
  // res.json() throws a cryptic "Unexpected token '<'" that bubbles to
  // the browser as a confusing JSON-parse error. Catch that shape and
  // raise something the user can actually act on.
  if (looksHtml) {
    // Strip tags/scripts/styles so the actual Google error message is
    // legible in the logs — this is the piece that tells us WHY the
    // response wasn't JSON (authorization page, "exceeded maximum
    // execution", a googleusercontent 404, etc.).
    const readable = text
      .replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<script[\s\S]*?<\/script>/gi, ' ')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&')
      .replace(/\s+/g, ' ')
      .trim()
    const titleMatch = text.match(/<title[^>]*>([\s\S]*?)<\/title>/i)
    console.error(
      `[AS] HTML-RESPONSE action=${action} status=${res.status} ` +
      `elapsed=${elapsed}ms ttfb=${ttfbMs}ms body=${bodyMs}ms ` +
      `elMax=${elMaxMs}ms zips=${batchDownloadsActive} ` +
      `respBytes=${text.length} finalUrl=${res.url}\n` +
      `[AS]   verdict=${elMaxMs > 1000 ? 'node-blocked (event loop stalled)' : 'apps-script-slow'}\n` +
      `[AS]   title="${titleMatch ? titleMatch[1].trim() : '(none)'}"\n` +
      `[AS]   readable="${readable.slice(0, 600)}"`
    )
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
    console.error(
      `[AS] NON-JSON action=${action} status=${res.status} respBytes=${text.length} ` +
      `head="${text.slice(0, 200).replace(/\n/g, ' ')}"`
    )
    throw new Error(
      `Apps Script response wasn't JSON (action="${action}", status=${res.status}): ` +
      `${text.slice(0, 200)}`
    )
  }
  if (json.error) {
    console.error(`[AS] JSON-ERROR action=${action} status=${res.status} error="${json.error}"`)
    throw new Error(json.error)
  }
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
 * Returns full WO list + stats for the admin dashboard. Decorates each
 * WO row with `invoice_view_url` (built from `qb_invoice_id`) so the
 * React side doesn't need to know about sandbox-vs-prod QBO base URLs.
 */
app.get('/api/dashboard', async (_req, res) => {
  try {
    const data = await callAppsScript('get_dashboard_data')
    if (Array.isArray(data?.wos)) {
      for (const wo of data.wos) {
        wo.invoice_view_url = buildInvoiceViewUrl(wo.qb_invoice_id)
      }
    }
    res.json(data)
  } catch (err) {
    console.error('GET /api/dashboard error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/revenue?start=YYYY-MM-DD&end=YYYY-MM-DD
 * Returns aggregated revenue for the Revenue dashboard tab. Both
 * dates inclusive; either can be omitted (server defaults to MTD).
 */
app.get('/api/revenue', async (req, res) => {
  try {
    const { start, end } = req.query
    const data = await callAppsScript('get_revenue_data', {
      start: start ? String(start) : '',
      end:   end   ? String(end)   : '',
    })
    res.json(data)
  } catch (err) {
    console.error('GET /api/revenue error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/production?start=YYYY-MM-DD&end=YYYY-MM-DD
 * Returns aggregated SF / LF / EA totals for the Production dashboard
 * tab. Same range conventions as /api/revenue.
 */
app.get('/api/production', async (req, res) => {
  try {
    const { start, end } = req.query
    const data = await callAppsScript('get_production_data', {
      start: start ? String(start) : '',
      end:   end   ? String(end)   : '',
    })
    res.json(data)
  } catch (err) {
    console.error('GET /api/production error:', err.message)
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
 * POST /api/field-report/check-shift-attribution
 * Pre-submit gate: warns the FR form when a Crew Chief appears to
 * still be wrapping up last night's overnight shift but the date
 * picker is set to today. Returns { should_confirm, prior_date,
 * reason }. The form shows a soft-warn modal on should_confirm:true.
 * Body: { crew_chief, work_date: 'YYYY-MM-DD' }
 */
app.post('/api/field-report/check-shift-attribution', async (req, res) => {
  try {
    const data = await callAppsScript('check_fr_shift_attribution', req.body)
    res.json(data)
  } catch (err) {
    console.error('POST /api/field-report/check-shift-attribution error:', err.message)
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
 * POST /api/tools/generate-pl-for-doc
 * Per-doc Production Log generation — fired from a Doc Status tab
 * pending-list "Generate" button. Apps Script validates SI Done for
 * the (date, contract, borough) before running. Body: { doc_id }.
 */
app.post('/api/tools/generate-pl-for-doc', async (req, res) => {
  try {
    const { doc_id } = req.body || {}
    if (!doc_id) return res.status(400).json({ error: 'Missing doc_id' })
    const data = await callAppsScript('generate_pl_for_doc', { doc_id })
    if (data && data.error) return res.status(400).json(data)
    res.json(data)
  } catch (err) {
    console.error('POST /api/tools/generate-pl-for-doc error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/tools/generate-cp-for-doc
 * Per-doc Certified Payroll generation. Validates that every worked
 * day's SI for the (week, contract, borough) is Done before running.
 * Body: { doc_id }.
 */
app.post('/api/tools/generate-cp-for-doc', async (req, res) => {
  try {
    const { doc_id, paystub } = req.body || {}
    if (!doc_id) return res.status(400).json({ error: 'Missing doc_id' })
    // `paystub` (optional) = { employees: [{ name, gross_pay, net_pay, deductions }] }
    // from the Upload Paystub step; Apps Script uses it to auto-fill
    // Withholdings & Net Pay and to run the gross-pay verification.
    const data = await callAppsScript('generate_cp_for_doc', { doc_id, paystub })
    if (data && data.error) return res.status(400).json(data)
    res.json(data)
  } catch (err) {
    console.error('POST /api/tools/generate-cp-for-doc error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/tools/generate-month-end-doc  { doc_id }
 * Fills a month-end doc (Employee Utilization / Certificates) and streams
 * the PDF back as a download — the user prints & signs by hand. Deliberately
 * NOT routed through Approvals. Apps Script supplies the field values
 * (build_month_end_fill); the Python worker's fill endpoint renders the PDF
 * with the proven pypdf + PyMuPDF process (correct /AP, still editable) so
 * we reuse the exact fill pipeline the other doc types use.
 *
 * Env: WORKER_FILL_URL (worker service base URL), FILL_SERVER_KEY (shared).
 */
app.post('/api/tools/generate-month-end-doc', async (req, res) => {
  try {
    const { doc_id } = req.body || {}
    if (!doc_id) return res.status(400).json({ error: 'Missing doc_id' })

    const spec = await callAppsScript('build_month_end_fill', { doc_id })
    if (!spec || spec.error) return res.status(400).json({ error: (spec && spec.error) || 'Could not build fill spec' })

    const fillBase = process.env.WORKER_FILL_URL
    if (!fillBase) return res.status(500).json({ error: 'WORKER_FILL_URL not configured on this server' })

    const wr = await fetch(fillBase.replace(/\/$/, '') + '/fill', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json', 'X-Fill-Key': process.env.FILL_SERVER_KEY || '' },
      body:    JSON.stringify({ doc_kind: spec.doc_kind, fields: spec.fields, filename: spec.filename }),
    })
    if (!wr.ok) {
      const t = await wr.text().catch(() => '')
      throw new Error(`Fill service error ${wr.status}: ${t.slice(0, 200)}`)
    }
    const out = Buffer.from(await wr.arrayBuffer())

    const safeName = String(spec.filename || 'month-end.pdf').replace(/[\r\n"]/g, '')
    res.setHeader('Content-Type', 'application/pdf')
    res.setHeader('Content-Disposition', `attachment; filename="${safeName}"`)
    res.send(out)
  } catch (err) {
    console.error('POST /api/tools/generate-month-end-doc error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/tools/generate-month-end-all  { doc_ids }
 * "Generate All" for a month: fills every given month-end doc, merges all
 * EUs into one PDF and all Certs into one PDF (collision-safe, in the
 * worker), and streams a single ZIP of the combined PDFs for download.
 */
app.post('/api/tools/generate-month-end-all', async (req, res) => {
  try {
    const { doc_ids } = req.body || {}
    if (!Array.isArray(doc_ids) || doc_ids.length === 0) return res.status(400).json({ error: 'No doc_ids' })

    const spec = await callAppsScript('build_month_end_fill_all', { doc_ids })
    if (!spec || spec.error) return res.status(400).json({ error: (spec && spec.error) || 'Could not build fill spec' })
    const groups = Array.isArray(spec.groups) ? spec.groups : []
    if (groups.length === 0) return res.status(400).json({ error: 'Nothing outstanding to generate' })

    const fillBase = process.env.WORKER_FILL_URL
    if (!fillBase) return res.status(500).json({ error: 'WORKER_FILL_URL not configured on this server' })

    // Fill + merge each group (EU, CERT) into one combined PDF via the worker.
    const files = []
    for (const g of groups) {
      const wr = await fetch(fillBase.replace(/\/$/, '') + '/fill-batch', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json', 'X-Fill-Key': process.env.FILL_SERVER_KEY || '' },
        body:    JSON.stringify({ doc_kind: g.doc_kind, items: g.items, combined_filename: g.combined_filename }),
      })
      if (!wr.ok) {
        const t = await wr.text().catch(() => '')
        throw new Error(`Fill service error ${wr.status}: ${t.slice(0, 200)}`)
      }
      files.push({ name: g.combined_filename, buf: Buffer.from(await wr.arrayBuffer()) })
    }

    const zipName = String(spec.zip_name || 'Month_End_Docs.zip').replace(/[\r\n"]/g, '')
    res.setHeader('Content-Type', 'application/zip')
    res.setHeader('Content-Disposition', `attachment; filename="${zipName}"`)
    const archive = archiver('zip', { zlib: { level: 0 } })   // PDFs already compressed
    archive.on('warning', (e) => console.warn('archiver warning:', e.message))
    archive.on('error', (e) => { console.error('archiver error:', e.message); try { res.end() } catch (_) {} })
    archive.pipe(res)
    for (const f of files) archive.append(f.buf, { name: f.name })
    await archive.finalize()
  } catch (err) {
    console.error('POST /api/tools/generate-month-end-all error:', err.message)
    if (!res.headersSent) res.status(500).json({ error: err.message })
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
 * POST /api/wo/:woId/status
 * Admin-driven WO status change from the Dashboard kebab. Bypasses the
 * one-way state machine in handleSubmitFieldReport_; admin can set any
 * status (including new "Returned" value). Audit row written to
 * Automation Log on success.
 *
 * Body: { status: 'Received' | 'Dispatched' | 'In Progress' | 'Completed' | 'Returned' }
 */
app.post('/api/wo/:woId/status', async (req, res) => {
  try {
    const data = await callAppsScript('update_wo_status', {
      wo_id:  req.params.woId,
      status: String(req.body?.status || '').trim(),
    })
    res.json(data)
  } catch (err) {
    console.error('POST /api/wo/:woId/status error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * DELETE /api/wo/:woId
 * Hard-deletes WO Tracker row + all Marking Items rows keyed to the WO.
 * Preserves Work Day Log / Sign-In Data / Doc Lifecycle Log / Drive
 * archive folder (audit trail). Logs to Automation Log. Returns
 * { marking_items_deleted, work_day_log_preserved } so the UI can show
 * a useful success toast.
 */
app.delete('/api/wo/:woId', async (req, res) => {
  try {
    const data = await callAppsScript('delete_wo', { wo_id: req.params.woId })
    res.json(data)
  } catch (err) {
    console.error('DELETE /api/wo/:woId error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/wo/:woId/edit-completed
 * Admin update to a Completed WO. Body:
 *   {
 *     as_of_date,
 *     issues,
 *     regen_mode: 'data_only' | 'replace_cfr' | 'new_cfr',
 *     include_in_production: boolean,
 *     production_date: 'YYYY-MM-DD'
 *   }
 * The marking-item edits themselves are already persisted via the
 * normal /api/marking-items PATCH endpoint with `preserve_completion`.
 */
app.post('/api/wo/:woId/edit-completed', async (req, res) => {
  try {
    const data = await callAppsScript('edit_completed_wo', {
      wo_id:                  req.params.woId,
      as_of_date:             req.body?.as_of_date || '',
      issues:                 req.body?.issues || '',
      regen_mode:             req.body?.regen_mode || 'data_only',
      include_in_production:  !!req.body?.include_in_production,
      production_date:        req.body?.production_date || '',
      marking_edits:          Array.isArray(req.body?.marking_edits)   ? req.body.marking_edits   : [],
      marking_adds:           Array.isArray(req.body?.marking_adds)    ? req.body.marking_adds    : [],
      marking_deletes:        Array.isArray(req.body?.marking_deletes) ? req.body.marking_deletes : [],
    })
    res.json(data)
  } catch (err) {
    console.error('POST /api/wo/:woId/edit-completed error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/wo/:woId/coordinates
 * Admin manual lat/lng entry from the Nav tab — fixes flagged
 * geocode warnings + provides an escape hatch when the geocoder
 * picked the wrong intersection. Body: { lat, lng }.
 */
app.post('/api/wo/:woId/coordinates', async (req, res) => {
  try {
    const data = await callAppsScript('update_wo_coordinates', {
      wo_id: req.params.woId,
      lat:   req.body?.lat,
      lng:   req.body?.lng,
    })
    res.json(data)
  } catch (err) {
    console.error('POST /api/wo/:woId/coordinates error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/wos/map
 * Returns the active WO set (Received / Dispatched / In Progress) for
 * the Nav tab map. Split into { mapped, unmapped } — unmapped are
 * surfaced in the "Needs coords" panel.
 */
app.get('/api/wos/map', async (_req, res) => {
  try {
    const data = await callAppsScript('get_wo_map_data')
    res.json(data)
  } catch (err) {
    console.error('GET /api/wos/map error:', err.message)
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
 * GET /api/pending-counts
 * Cold-start endpoint for the webapp's nav badges. Returns three cheap
 * counts in one round-trip: docs needing review, approved-docs waiting
 * for the worker, and pending sign-ins. Doc Status pending is NOT here
 * — that's expensive and is populated from the regular /api/doc-status
 * fetch when the user actually visits Dashboard.
 */
app.get('/api/pending-counts', async (_req, res) => {
  try {
    const data = await callAppsScript('get_pending_counts')
    res.json(data)
  } catch (err) {
    console.error('GET /api/pending-counts error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/pending-counts/doc-status
 * Heavier sibling of /api/pending-counts that builds the Doc Status
 * payload to derive its pending count. Fired in parallel with the
 * cheap counts so nav badges don't have to wait on it.
 */
app.get('/api/pending-counts/doc-status', async (_req, res) => {
  try {
    const data = await callAppsScript('get_doc_status_pending_count')
    res.json(data)
  } catch (err) {
    console.error('GET /api/pending-counts/doc-status error:', err.message)
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
 * POST /api/approvals/:fileId/reupload
 * Replaces a pending-approval PDF (sitting in a Docs Needing Review
 * subfolder) with a new signed/rescanned version. Multipart field:
 *   file — the replacement PDF (raw scan or chosen PDF; no OCR/parse).
 * Apps Script recreates the file with the SAME name in the SAME folder
 * and trashes the original, so a NEW file_id comes back.
 */
app.post('/api/approvals/:fileId/reupload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file attached' })
    const base64 = req.file.buffer.toString('base64')
    const data = await callAppsScript('reupload_pending_approval', {
      file_id:   req.params.fileId,
      bytes_b64: base64,
    })
    if (data && data.error) return res.status(500).json(data)
    res.json(data)   // { success, file_id, filename }
  } catch (err) {
    console.error(`POST /api/approvals/${req.params.fileId}/reupload error:`, err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/approvals/:fileId/signin-rows?filename=...
 * Returns the Daily Sign-In Data rows for a pending sign-in sheet so the
 * Approvals page can show/edit the recorded hours. `filename` (carried
 * by the approvals list item) is the lookup key — it encodes contract /
 * borough / date / crew chief. Response:
 *   { rows: [{ row_index, name, classification, time_in, time_out,
 *              hours, overtime, crew_chief }],
 *     other_hours: { name: hours },     // other sheets, same day
 *     meta: { date, contract, borough, chief_slug, ambiguous } }
 */
app.get('/api/approvals/:fileId/signin-rows', async (req, res) => {
  try {
    const data = await callAppsScript('list_signin_rows_for_file', {
      file_id:  req.params.fileId,
      filename: req.query.filename || '',
    })
    res.json(data)
  } catch (err) {
    console.error(`GET /api/approvals/${req.params.fileId}/signin-rows error:`, err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/approvals/:fileId/signin-header
 * Read-only header metadata for a submitted sign-in — the fields the
 * Approvals-page header card mirrors from the Sign-In tab (contract,
 * billing identity, prime contractor, address, work orders, crew chief).
 * Response: { header: { date, contract_number, borough,
 *   bill_contract_number, bill_borough, contractor, prime_contractor,
 *   subcontractor, address, crew_chief, wos: [{id, location}] } }
 */
app.get('/api/approvals/:fileId/signin-header', async (req, res) => {
  try {
    const data = await callAppsScript('signin_header_for_file', {
      file_id:  req.params.fileId,
      filename: req.query.filename || '',
    })
    res.json(data)
  } catch (err) {
    console.error(`GET /api/approvals/${req.params.fileId}/signin-header error:`, err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/approvals/:fileId/save-signin-rows
 * Admin edits to a submitted sign-in's hours. Writes Classification /
 * Time In / Time Out / Hours back to Daily Sign-In Data and recomputes
 * Overtime for the whole day. The signed PDF is left untouched (the
 * sheet is the payroll source of truth). Body: { filename, rows }.
 */
app.post('/api/approvals/:fileId/save-signin-rows', async (req, res) => {
  try {
    const data = await callAppsScript('save_signin_row_edits', {
      file_id:  req.params.fileId,
      filename: req.body.filename || '',
      rows:     req.body.rows || [],
    })
    res.json(data)
  } catch (err) {
    console.error(`POST /api/approvals/${req.params.fileId}/save-signin-rows error:`, err.message)
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
    //
    // We do NOT flatten and do NOT use form.setText() for the principal
    // sign-off fields. Both of those paths cause pdf-lib to regenerate
    // /AP streams using its own font/sizing logic, which destroys any
    // pypdf-rendered content already on the form (the crew-leader rows).
    //
    // Instead we draw the principal name + title + date + signature
    // image directly onto the page content stream and save without
    // touching appearances. The PDF arrives from the worker with
    // /NeedAppearances=false and correct PyMuPDF-rendered /AP streams
    // already in place — pdf-lib's save({ updateFieldAppearances: false })
    // preserves both, so no extra setup is needed here.
    const { PDFDocument, StandardFonts } = await import('pdf-lib')
    const pdfDoc = await PDFDocument.load(pdfBytes, { updateMetadata: false })
    const form   = pdfDoc.getForm()
    const helv   = await pdfDoc.embedFont(StandardFonts.Helvetica)

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

    // Helper: locate a field's widget rect + page.
    const locateField = (fieldName) => {
      const field   = form.getField(fieldName)
      const widget  = field.acroField.getWidgets()[0]
      const rect    = widget.getRectangle()
      const pageRef = widget.P()
      const pages   = pdfDoc.getPages()
      const page    = pages.find(p => p.ref === pageRef) || pages[0]
      return { rect, page }
    }

    // Draw centered text on the page directly (no setText — that would
    // force /AP regeneration on flatten/save).
    const drawCenteredText = (fieldName, text, fontSize) => {
      try {
        const { rect, page } = locateField(fieldName)
        const w = helv.widthOfTextAtSize(text, fontSize)
        const h = helv.heightAtSize(fontSize)
        const x = rect.x + (rect.width  - w) / 2
        const y = rect.y + (rect.height - h) / 2
        page.drawText(text, { x, y, size: fontSize, font: helv })
      } catch (e) {
        console.warn(`drawCenteredText failed on ${fieldName}: ${e.message}`)
      }
    }
    drawCenteredText('Contractor_Name',      String(name),  11)
    drawCenteredText('Contractor_Title',     String(title), 11)
    drawCenteredText('Date_Signature_Block', dateStr,       11)

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
      const { rect, page } = locateField('Contractor_Signature')

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
      // Fall through — name/title/date are still applied via drawText
    }

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
 * POST /api/approvals/:fileId/approve-cert-payroll
 * Certified-Payroll-specific approval: same modal-driven flow as
 * /approve-signin (signature + printed name + title from
 * PrincipalSignModal), but targets the CP form's field names and
 * the Signature_Block rect added in Acrobat.
 *
 * Date format on the CP signature line is "<Month> <D>, 20YY"
 * (e.g. "May 2, 2026"). The form has the literal "20" pre-printed
 * between the DATE and YEAR fields, so:
 *   - DATE field gets "<Full Month> <Day>,"   e.g. "May 2,"
 *   - YEAR field gets the 2-digit year         e.g. "26"
 *
 * Body: { signature_b64: string, name: string, title: string }
 */
app.post('/api/approvals/:fileId/approve-cert-payroll', express.json({ limit: '5mb' }), async (req, res) => {
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
    //
    // CP has many pypdf-rendered worker rows already on the form. We
    // must not let pdf-lib regenerate /AP streams (default behavior of
    // form.flatten() and form.setText()) — that would re-render every
    // worker field with pdf-lib's default font/sizing and destroy the
    // formatting (truncated addresses, missing checkboxes, etc).
    //
    // Same approach as the sign-in route: draw the four signature-block
    // values + the signature image directly onto the page content
    // stream and save without touching appearances. The PDF arrives
    // from the worker with /NeedAppearances=false and correct PyMuPDF-
    // rendered /AP streams already in place; pdf-lib preserves both.
    const { PDFDocument, StandardFonts } = await import('pdf-lib')
    const pdfDoc = await PDFDocument.load(pdfBytes, { updateMetadata: false })
    const form   = pdfDoc.getForm()
    const helv   = await pdfDoc.embedFont(StandardFonts.Helvetica)

    // Date: today in America/New_York. Same TZ rationale as the sign-in
    // route — without the override Node runs in UTC and rolls over at 20:00 ET.
    const parts = new Intl.DateTimeFormat('en-US', {
      timeZone:  'America/New_York',
      month:     'long',
      day:       'numeric',
      year:      '2-digit',
    }).formatToParts(new Date())
    const pick = (t) => parts.find(p => p.type === t)?.value || ''
    const dateStr = `${pick('month')} ${pick('day')},`
    const yearStr = pick('year')

    // Helper: locate a field's widget rect + page.
    const locateField = (fieldName) => {
      const field   = form.getField(fieldName)
      const widget  = field.acroField.getWidgets()[0]
      const rect    = widget.getRectangle()
      const pageRef = widget.P()
      const pages   = pdfDoc.getPages()
      const page    = pages.find(p => p.ref === pageRef) || pages[0]
      return { rect, page }
    }

    // Draw centered text on the page directly (no setText — that would
    // force /AP regeneration on flatten/save). 12pt Helvetica fits the
    // CP signature-block field rects (~29pt tall) with comfortable padding.
    const SIG_TEXT_SIZE = 12
    const drawCenteredText = (fieldName, text) => {
      try {
        const { rect, page } = locateField(fieldName)
        const w = helv.widthOfTextAtSize(text, SIG_TEXT_SIZE)
        const h = helv.heightAtSize(SIG_TEXT_SIZE)
        const x = rect.x + (rect.width  - w) / 2
        const y = rect.y + (rect.height - h) / 2
        page.drawText(text, { x, y, size: SIG_TEXT_SIZE, font: helv })
      } catch (e) {
        console.warn(`drawCenteredText failed on ${fieldName}: ${e.message}`)
      }
    }
    drawCenteredText('OFFICER OR PRINCIPAL (print)', String(name))
    drawCenteredText('TITLE',                        String(title))
    drawCenteredText('DATE',                         dateStr)
    drawCenteredText('YEAR',                         yearStr)

    // Overlay the signature PNG onto the Signature_Block field's rect.
    // Same padding + scaling approach as the sign-in route.
    const SIG_H_PAD = 1.0
    const SIG_V_PAD = 3.0
    const pngB64 = String(signature_b64).replace(/^data:image\/png;base64,/, '')
    const pngBytes = Buffer.from(pngB64, 'base64')
    const sigImage = await pdfDoc.embedPng(pngBytes)
    try {
      const { rect, page } = locateField('Signature_Block')

      const boxX = rect.x      - SIG_H_PAD
      const boxY = rect.y      - SIG_V_PAD
      const boxW = rect.width  + 2 * SIG_H_PAD
      const boxH = rect.height + 2 * SIG_V_PAD

      const scale = Math.min(boxW / sigImage.width, boxH / sigImage.height)
      const drawW = sigImage.width  * scale
      const drawH = sigImage.height * scale
      const drawX = boxX + (boxW - drawW) / 2
      const drawY = boxY + (boxH - drawH) / 2
      page.drawImage(sigImage, { x: drawX, y: drawY, width: drawW, height: drawH })
    } catch (e) {
      console.warn('pdf-lib signature overlay failed:', e.message)
      // Fall through — name/title/date are still applied via drawText
    }

    const signedBytes = await pdfDoc.save({ updateFieldAppearances: false })

    // 3. Upload patched bytes to Approved Docs + trash original via Apps Script.
    // Reuses approve_signin_with_bytes — that handler is doc-type
    // agnostic (just swaps bytes and routes to Approved Docs), so it
    // serves both sign-ins and certified payrolls.
    const result = await callAppsScript('approve_signin_with_bytes', {
      file_id:   fileId,
      filename:  filename,
      bytes_b64: Buffer.from(signedBytes).toString('base64'),
    })
    if (result.error) throw new Error(result.error)
    res.json(result)
  } catch (err) {
    console.error(`POST /api/approvals/${fileId}/approve-cert-payroll error:`, err.message)
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
 * POST /api/tools/paystub/parse
 * Reads a payroll paystub with Claude vision so the CP Generate modal can
 * auto-fill each employee's Withholdings & Deductions and Net Pay (today
 * transcribed by hand from payrollforconstruction.com). Standard input is
 * the "Pre-Check Register" (one page, every employee, two rows each), as a
 * PDF or a JPEG/PNG screenshot; individual check stubs are also handled.
 *
 * Employee names are resolved against the Employee Registry closed list
 * (same closed-list trick as /api/signin-queue/parse-upload) so the "Last,
 * First M" register spelling maps to the exact registry name the CP data
 * joins on.
 *
 * Returns { employees: [{ name, employee_number, gross_pay, net_pay,
 * deductions }], filename }. No persistence — the caller passes the result
 * straight into POST /api/tools/generate-cp-for-doc.
 */
app.post('/api/tools/paystub/parse', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file attached' })
    if (!process.env.ANTHROPIC_API_KEY) {
      return res.status(500).json({ error: 'ANTHROPIC_API_KEY not configured on this server' })
    }
    const mime = req.file.mimetype || 'application/pdf'
    const ALLOWED = ['application/pdf', 'image/jpeg', 'image/png']
    if (!ALLOWED.includes(mime)) {
      return res.status(400).json({ error: 'Unsupported file type (got ' + mime + '). Upload a PDF, JPEG, or PNG.' })
    }
    const base64 = req.file.buffer.toString('base64')
    // PDFs go in a `document` block; images in an `image` block.
    const sourceBlock = mime === 'application/pdf'
      ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: base64 } }
      : { type: 'image',    source: { type: 'base64', media_type: mime,              data: base64 } }

    // Closed employee list for exact name resolution (see parse-upload).
    let employeeList = []
    try {
      const empRes = await callAppsScript('list_employees')
      employeeList = (empRes && Array.isArray(empRes.employees))
        ? empRes.employees.map(e => String(e.name || '').trim()).filter(Boolean)
        : []
    } catch (e) {
      console.warn('Paystub upload: could not fetch employee list:', e.message)
    }
    const employeeListText = employeeList.length
      ? employeeList.map(n => `  - ${n}`).join('\n')
      : '  (registry currently empty — return the name as printed)'

    const prompt = `You are extracting per-employee pay figures from a construction payroll paystub produced by payrollforconstruction.com. Return ONLY valid JSON — no markdown fences, no commentary.

The document is USUALLY a "Pre-Check Register": a single table listing every employee for one weekly pay period. Each employee occupies TWO stacked rows:
  - Row 1 columns: REG (hours) | QTY | REG Wages | Taxable Add | Tx Uni Frng | Gross Pay | FICA | State | Union | Net Pay
  - Row 2 columns: OVT (hours) | OTH | OVT Wages | Non-Tax Add | Emp Fringe | Tot Taxable | Federal | Local | Misc | Pay Method
Each employee block is preceded by an employee number and a name like "Angelides , Stamatis D".

For EACH employee return:
  - gross_pay: the number in the "Gross Pay" column (row 1) — the full weekly gross across all work.
  - net_pay:   the number in the "Net Pay" column (row 1).
  - deductions: gross_pay minus net_pay. (Sanity-check: it should also equal FICA + State + Union + Federal + Local + Misc summed across both rows — if that sum disagrees, still return gross_pay - net_pay.)
  - employee_number: the payroll ID printed with the block (e.g. "3540"), or null.

If instead the document is an individual check stub (one employee per page, with a "Summary" block), read Gross Pay, Total Deductions, and Net Pay directly from that Summary; deductions = Total Deductions.

IGNORE any totals/summary rows at the bottom of a register (rows labeled "Employees:", "Total Checks", "Total Direct Deposits", "Total Adjustments"). Only return real employees.

Schema:
{
  "employees": [
    {
      "name":            "<EXACTLY one name from the Employee Registry below — see Name matching rules>",
      "employee_number": "<the payroll ID printed for this employee, or null>",
      "gross_pay":       <number, no $ or commas>,
      "net_pay":         <number, no $ or commas>,
      "deductions":      <number, no $ or commas>
    }
  ]
}

Employee Registry (the ONLY valid values for employees[].name):
${employeeListText}

Name matching rules:
- The register prints names "Last , First M". For each, find the SINGLE registry entry above whose person matches (account for last/first order, middle initials, accents, abbreviations, minor misspellings — pick the best fit).
- Return that registry entry verbatim — exact capitalization and spelling as listed above.
- If the registry list is empty, return the name as printed on the paystub.
- Skip a row only if it has no employee at all.

Numbers:
- Strip "$" and thousands separators; "1,529.55" → 1529.55.
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
        max_tokens: 4096,
        messages: [{
          role: 'user',
          content: [ sourceBlock, { type: 'text', text: prompt } ],
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

    let raw = String(textBlock.text || '').trim()
    raw = raw.replace(/^```(?:json)?\s*/i, '').replace(/```$/, '').trim()
    let parsed
    try {
      parsed = JSON.parse(raw)
    } catch (e) {
      throw new Error('Claude returned non-JSON text — first 200 chars: ' + raw.slice(0, 200))
    }

    // Normalize into a clean numeric array; drop rows missing a name.
    const num = v => {
      const n = Number(String(v == null ? '' : v).replace(/[^0-9.\-]/g, ''))
      return isFinite(n) ? n : null
    }
    const employees = (Array.isArray(parsed.employees) ? parsed.employees : [])
      .map(e => ({
        name:            String(e.name || '').trim(),
        employee_number: e.employee_number != null ? String(e.employee_number).trim() : null,
        gross_pay:       num(e.gross_pay),
        net_pay:         num(e.net_pay),
        deductions:      num(e.deductions),
      }))
      .filter(e => e.name)

    res.json({ employees, filename: req.file.originalname })
  } catch (err) {
    console.error('POST /api/tools/paystub/parse error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/metro-delivery/list
 * One-off bundling helper. Returns metadata for every Metro Express
 * Status='Completed' WO's archived PDF + every Metro production log.
 * Used to compile a manual Metro delivery batch.
 */
app.get('/api/metro-delivery/list', async (_req, res) => {
  try {
    const data = await callAppsScript('list_metro_completed_docs')
    res.json(data)
  } catch (err) {
    console.error('GET /api/metro-delivery/list error:', err.message)
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
  // Timing: most photo-upload slowness is "where is the second spent".
  // Log each leg so Railway logs show it: total = encode + AS round-trip.
  const t0 = Date.now()
  try {
    if (!req.file)        return res.status(400).json({ error: 'No file attached' })
    if (!req.body.wo_id)  return res.status(400).json({ error: 'wo_id is required' })

    const sizeKB = (req.file.size / 1024).toFixed(0)
    const base64 = req.file.buffer.toString('base64')
    const t1 = Date.now()
    const data   = await callAppsScript('upload_photo', {
      wo_id:     req.body.wo_id,
      filename:  req.file.originalname,
      mime_type: req.file.mimetype,
      data:      base64
    })
    const t2 = Date.now()
    console.log(`[upload-photo] wo=${req.body.wo_id} size=${sizeKB}KB ` +
                `encode=${t1 - t0}ms appsScript=${t2 - t1}ms total=${t2 - t0}ms`)
    res.set('Server-Timing', `encode;dur=${t1 - t0}, appsScript;dur=${t2 - t1}, total;dur=${t2 - t0}`)
    res.json({ ...(data || {}), _server_ms: t2 - t0, _appsScript_ms: t2 - t1 })
  } catch (err) {
    console.error(`POST /api/upload-photo error after ${Date.now() - t0}ms:`, err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/wo-photos/:woId
 * Lists the photos currently in the WO's Drive Photos/ folder so the
 * Field Report page can render a "previously taken" gallery.
 */
app.get('/api/wo-photos/:woId', async (req, res) => {
  try {
    const data = await callAppsScript('list_wo_photos', { wo_id: req.params.woId })
    res.json(data)
  } catch (err) {
    console.error('GET /api/wo-photos error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/wo-photos/:fileId/content
 * Streams the raw bytes of a Drive photo back as image/*. The webapp
 * lightbox uses this to render historic photos at full size — their
 * list_wo_photos thumbnail is only 220x220 and unreadable for review.
 */
app.get('/api/wo-photos/:fileId/content', async (req, res) => {
  try {
    const data = await callAppsScript('get_wo_photo_content', { file_id: req.params.fileId })
    if (data.error) return res.status(500).json({ error: data.error })
    const buf = Buffer.from(data.data || '', 'base64')
    res.set('Content-Type', data.mime || 'image/jpeg')
    res.set('Cache-Control', 'private, max-age=3600')
    res.send(buf)
  } catch (err) {
    console.error('GET /api/wo-photos/:fileId/content error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * DELETE /api/wo-photos/:fileId
 * Trashes a photo from Drive when the user removes it from the gallery.
 * Reuses the existing `trash_file` Apps Script action.
 */
app.delete('/api/wo-photos/:fileId', async (req, res) => {
  try {
    const data = await callAppsScript('trash_file', { file_id: req.params.fileId })
    res.json(data)
  } catch (err) {
    console.error('DELETE /api/wo-photos error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/reverse-geocode
 * Reverse-geocodes a lat/lng via Apps Script's built-in Maps service.
 * Used by the photo watermark pipeline.
 * Body: { lat, lng }
 */
app.post('/api/reverse-geocode', async (req, res) => {
  try {
    const data = await callAppsScript('reverse_geocode', req.body)
    res.json(data)
  } catch (err) {
    console.error('POST /api/reverse-geocode error:', err.message)
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


/**
 * POST /api/documents/flags
 * Manual Done/Sent toggles from the DocStatusChips component on the
 * WO Tracker tab. Body matches set_docs_sent verbatim:
 *   { updates: [{ wo_id, doc_type, done?, sent? }] }
 * Apps Script accepts both friendly (CFR) and internal (Field Report)
 * doc_type names and bumps the dashboard cache after writes.
 */
app.post('/api/documents/flags', async (req, res) => {
  try {
    const data = await callAppsScript('set_docs_sent', req.body)
    res.json(data)
  } catch (err) {
    console.error('POST /api/documents/flags error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * GET /api/doc-status?month=YYYY-MM
 * Returns the calendar payload for the Doc Status tab. Defaults to
 * current month when month is missing.
 */
app.get('/api/doc-status', async (req, res) => {
  try {
    const month = String(req.query?.month || '').trim()
    const data = await callAppsScript('get_doc_status_calendar', { month })
    res.json(data)
  } catch (err) {
    console.error('GET /api/doc-status error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/doc-status/flags
 * Per-doc lifecycle update for time-anchored docs (PL / SI / CP).
 * Body: { updates: [{ doc_id, done?, sent? }, ...] }
 *
 * Companion to /api/documents/flags, which still handles per-WO
 * updates for CFR + Invoice. Sibling endpoints by design — the client
 * calls whichever one matches the identifier shape it has in hand.
 */
app.post('/api/doc-status/flags', async (req, res) => {
  try {
    const data = await callAppsScript('set_doc_status', req.body)
    res.json(data)
  } catch (err) {
    console.error('POST /api/doc-status/flags error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/documents/list-batch
 * Returns metadata for documents matching the given filters — used by
 * the DownloadDocumentsModal's preview step before the user commits to
 * the actual zip download. No file bytes returned. Body shape matches
 * the Apps Script list_documents_for_batch action.
 */
app.post('/api/documents/list-batch', async (req, res) => {
  try {
    const data = await callAppsScript('list_documents_for_batch', req.body)
    res.json(data)
  } catch (err) {
    console.error('POST /api/documents/list-batch error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

/**
 * POST /api/documents/batch-download
 * Streams a zip of every doc matching the filters to the client.
 * Body: { ...filters, mark_sent: bool }.
 *
 * Flow:
 *   1. Call Apps Script for the list of files.
 *   2. archiver streams the zip to res; MANIFEST.txt at the root.
 *   3. Each file's bytes fetched via Apps Script get_drive_file_bytes.
 *   4. On client abort (req.aborted / req.close), stop the file loop
 *      and skip mark_sent.
 *   5. On success and mark_sent=true, fire-and-forget set_docs_sent
 *      to flip Sent flags for every (wo_id, doc_type) covered.
 */
app.post('/api/documents/batch-download', express.json({ limit: '1mb' }), async (req, res) => {
  const filters  = req.body || {}
  const markSent = !!filters.mark_sent

  // Counted so the [AS] log lines can attribute a latency spike on some
  // unrelated endpoint to a zip stream running concurrently.
  batchDownloadsActive++
  res.on('close', () => { batchDownloadsActive-- })

  // Cancel detection: listen on RES (not req). req.close fires when the
  // incoming request stream ends, which for a POST happens once express
  // has parsed the body — that's BEFORE our handler runs. Using req
  // produced spurious "cancelled" flags that broke the zip stream and
  // skipped the post-finish set_docs_sent call. res.on('close') fires
  // when the response stream closes; combined with !res.writableEnded
  // it's the canonical "client disconnected mid-stream" signal.
  let cancelled = false
  res.on('close', () => {
    if (!res.writableEnded) {
      console.warn('batch-download: response closed before writableEnded — cancelled')
      cancelled = true
    }
  })

  try {
    // Resolve the file list first (small JSON payload).
    const listing = await callAppsScript('list_documents_for_batch', filters)
    const files = Array.isArray(listing.files) ? listing.files : []
    if (files.length === 0) {
      return res.status(400).json({
        error:    'No documents matched the requested filters.',
        missing:  listing.missing  || [],
        warnings: listing.warnings || [],
      })
    }
    console.log(`batch-download: starting zip with ${files.length} files`)

    const ts = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14)
    res.setHeader('Content-Type', 'application/zip')
    res.setHeader('Content-Disposition', `attachment; filename="oneiro-docs-${ts}.zip"`)
    res.setHeader('Cache-Control', 'no-store')
    res.flushHeaders()

    // level 0 (store), not 9. The entries are PDFs, which are already
    // deflate-compressed internally — level 9 bought ~nothing in size and
    // cost a lot of CPU. Worse, zlib runs on libuv's 4-thread pool, which
    // is the same pool Node's DNS getaddrinfo uses, so compressing a large
    // batch starved every outbound Apps Script fetch of a lookup thread.
    const archive = archiver('zip', { zlib: { level: 0 } })
    archive.on('warning', (e) => console.warn('archiver warning:', e.message))
    // Don't call res.end() in the error handler — archiver's pipe will
    // unpipe and let the response surface the error naturally. Manually
    // ending the response here used to truncate the stream while bytes
    // were still in flight.
    archive.on('error', (e) => {
      console.error('archiver error:', e && e.stack || e.message)
    })
    archive.pipe(res)

    // Build & append manifest first so it lives at the zip root.
    const manifest = buildBatchManifest(filters, listing)
    archive.append(manifest, { name: 'MANIFEST.txt' })

    // Determine zip-path layout: contractor folders only when more
    // than one contractor is present in the result set.
    const contractorsInBatch = new Set(files.map(f => f.contractor || 'Unknown'))
    const includeContractorFolder = contractorsInBatch.size > 1

    // Track filename collisions inside each (contractor)/(doc-type)
    // path; on collision prefix with the WO id.
    const docTypeFolder = (dt) => ({
      'CFR':               'Contractor Field Reports',
      'Production Log':    'Production Logs',
      'Sign-In':           'Sign-In Sheets',
      'Certified Payroll': 'Certified Payroll',
      'Invoice':           'Invoices',
    })[dt] || dt
    const usedPaths = new Set()
    const zipPathFor = (file) => {
      // Bundled entries (CP+SI bundle, photos) carry an explicit zip_path
      // from the listing handler — honor it verbatim. Defensive uniqueness
      // check below still applies.
      let candidate
      if (file.zip_path) {
        candidate = String(file.zip_path)
      } else {
        const folder = (includeContractorFolder ? `${file.contractor}/` : '') + docTypeFolder(file.doc_type)
        candidate = `${folder}/${file.filename}`
      }
      if (!usedPaths.has(candidate)) {
        usedPaths.add(candidate)
        return candidate
      }
      // Collision — prefix with WO id (or first wo id for multi-WO files)
      const prefix = (file.wo_ids && file.wo_ids[0]) || 'unknown-wo'
      const dir = candidate.substring(0, candidate.lastIndexOf('/'))
      const base = candidate.substring(candidate.lastIndexOf('/') + 1)
      candidate = `${dir}/${prefix}_${base}`
      let i = 2
      while (usedPaths.has(candidate)) {
        candidate = `${dir}/${prefix}_${i}_${base}`
        i++
      }
      usedPaths.add(candidate)
      return candidate
    }

    let appended = 0
    let failed   = 0
    for (const file of files) {
      if (cancelled) break
      try {
        const fetched = await callAppsScript('get_drive_file_bytes', { file_id: file.file_id })
        if (!fetched || !fetched.data) {
          console.warn(`batch-download: get_drive_file_bytes returned no bytes for ${file.filename}`)
          failed++
          continue
        }
        const buf = Buffer.from(fetched.data, 'base64')
        const prepared = await prepPdfForDelivery(buf, file.filename)
        archive.append(prepared, { name: zipPathFor(file) })
        appended++
      } catch (e) {
        failed++
        console.warn(`batch-download: failed to fetch ${file.filename}: ${e.message}`)
        // Continue with the other files rather than aborting the whole zip.
      }
    }
    console.log(`batch-download: appended=${appended} failed=${failed} cancelled=${cancelled}`)

    if (cancelled) {
      // Client gave up — abort the archive (which destroys the underlying
      // stream including the response) and skip set_docs_sent.
      try { archive.abort() } catch (_) {}
      return
    }

    // finalize() returns a Promise in archiver v5+. Awaiting it lets us
    // catch finalization errors and confirm bytes flushed before we
    // schedule the set_docs_sent call.
    try {
      await archive.finalize()
      console.log('batch-download: archive finalized')
    } catch (e) {
      console.error('batch-download: archive.finalize() threw:', e && e.stack || e.message)
      return
    }

    // Wait for the response stream to fully drain before flipping Sent
    // flags. res.on('finish') is the right signal — it fires after the
    // last byte was written to the socket.
    //
    // Two storage shapes to handle:
    //   - CFR + Invoice  → per-WO via /api/documents/flags-style updates
    //                      (wo_id + doc_type + sent).
    //   - PL / SI / CP   → per-doc via Doc Lifecycle Log keyed by Doc ID.
    //                      Doc ID = <PREFIX>_<ANCHOR>_<CONTRACTNUM>_<BOROUGH>
    //                      where ANCHOR is the work_date (already in
    //                      the listing payload) and PREFIX is PL/SI/CP.
    res.on('finish', () => {
      if (cancelled || !markSent) return

      const PER_WO = { CFR: 1, Invoice: 1 }
      // SI Sent isn't tracked (Sign-Ins ride out with the CP, CP Sent
      // implies SI Sent) so it's omitted here.
      // PL is keyed per-(date, contractor, crew_chief) — multi-crew shifts
      // get one row per chief. CP is per-(weekstart, contract, borough).
      // We use the listing's precomputed f.doc_id so the per-crew suffix
      // is honored rather than reconstructing a chief-less id here.

      const perWoUpdates = []
      const perDocUpdates = []
      const seenWo = new Set()
      const seenDoc = new Set()

      const slug = (s) => String(s || '').trim().replace(/\s+/g, '_')

      files.forEach(f => {
        // Bundled entries (CP+SI bundle, Photos) ride along with other
        // docs and don't carry their own Sent semantic — skip the flip.
        if (f.bundled) return
        if (PER_WO[f.doc_type]) {
          (f.wo_ids || []).forEach(woId => {
            const k = woId + '|' + f.doc_type
            if (seenWo.has(k)) return
            seenWo.add(k)
            perWoUpdates.push({ wo_id: woId, doc_type: f.doc_type, sent: true })
          })
          return
        }
        // Only PL + CP carry a tracked Sent flag (SI rides out with the CP).
        if (f.doc_type !== 'Production Log' && f.doc_type !== 'Certified Payroll') return
        // Prefer the lifecycle doc_id the listing handler computed — for
        // multi-crew PLs it already includes the per-crew `_chief-<slug>`
        // suffix, so each crew's PL flips its OWN row instead of colliding
        // on a chief-less id (which left them perpetually "unsent" and
        // re-downloaded). Fall back to reconstruction only if it's absent.
        let docId = String(f.doc_id || '').trim()
        if (!docId) {
          const anchor = String(f.work_date || '').trim()
          if (!anchor) return
          if (f.doc_type === 'Production Log') {
            const contractor = slug(f.contractor)
            if (!contractor) return
            docId = `PL_${anchor}_${contractor}`
          } else {
            const cn = String(f.contract_num || '').split('/')[0].trim()
            const borough = String(f.borough || '').trim()
            if (!cn || !borough) return
            docId = `CP_${anchor}_${cn}_${borough}`
          }
        }
        if (seenDoc.has(docId)) return
        seenDoc.add(docId)
        perDocUpdates.push({ doc_id: docId, sent: true })
      })

      if (perWoUpdates.length) {
        console.log(`batch-download: marking ${perWoUpdates.length} per-WO (CFR/INV) entries sent`)
        callAppsScript('set_docs_sent', { updates: perWoUpdates }).catch(e => {
          console.warn('batch-download: set_docs_sent (per-WO) failed:', e.message)
        })
      }
      if (perDocUpdates.length) {
        console.log(`batch-download: marking ${perDocUpdates.length} per-doc (PL/SI/CP) entries sent`)
        callAppsScript('set_doc_status', { updates: perDocUpdates }).catch(e => {
          console.warn('batch-download: set_doc_status (per-doc) failed:', e.message)
        })
      }
    })
  } catch (err) {
    console.error('POST /api/documents/batch-download error:', err && err.stack || err.message)
    if (!res.headersSent) {
      res.status(500).json({ error: err.message })
    }
    // If headers already went out, the partial zip the browser has is
    // unrecoverable — let the response close naturally rather than
    // forcing res.end() (which used to truncate good data in flight).
  }
})

/**
 * Prep a single PDF for inclusion in the delivery zip.
 *
 * Why: every CFR (and other doc-type) coming out of the worker uses
 * the same AcroForm field names from a shared template (Hours,
 * Quantity_1, Date, etc). When admin combines individual zip PDFs
 * via macOS's Create PDF Quick Action, the merged document has a
 * single AcroForm tree where the duplicated names collapse —
 * Acrobat shows only the first PDF's filled values, the rest appear
 * blank.
 *
 * Approach: rename every field's partial name with a per-file
 * suffix and mark each field read-only. Combine tools see distinct
 * names → no collision. Primes can't edit (read-only flag respected
 * by Acrobat / Preview). Crucially, we never touch /AP — the
 * PyMuPDF-rendered appearance is preserved byte-for-byte.
 *
 * This deliberately avoids pdf-lib's `form.flatten()`, which:
 *   - throws on certain merged WO docs (CFRs use _replaceArchivedWODoc_),
 *     causing silent fall-through to an editable original;
 *   - doesn't apply XObject /Matrix offsets when emitting /Do ops,
 *     producing visible misalignment on radio button selected-state
 *     indicators (Production Log borough circles).
 *
 * Drive originals stay untouched + editable; only the zip-bound copy
 * is renamed + locked.
 *
 * Falls back to the original buffer on:
 *   - non-PDF inputs (Invoice_*.txt artifacts, magic-byte check)
 *   - PDFs that provably carry no AcroForm (see canHaveAcroForm)
 *   - PDFs with no form fields (nothing to do)
 *   - pdf-lib parse errors (defensive)
 */
/**
 * Cheap, CONSERVATIVE pre-check: could this PDF possibly have an AcroForm?
 *
 * Returns false only when we can prove it cannot, so a false answer is
 * always safe to act on. Getting this backwards is dangerous: a PDF whose
 * fields we fail to rename keeps its generic field names, and when a prime
 * merges several of our docs the same-named fields collapse into one value
 * — the exact bug prepPdfForDelivery exists to prevent.
 *
 * Two literals matter:
 *   /AcroForm — present uncompressed when the catalog is a plain object.
 *   /ObjStm   — a compressed object stream, which CAN hide the catalog
 *               (and therefore /AcroForm) from a raw byte scan. PDF 1.5+.
 *
 * An object stream's own dictionary is never itself compressed, so /ObjStm
 * always appears in the clear. If NEITHER literal is present, there is no
 * object stream to hide a catalog in and no visible /AcroForm — hence no
 * form.
 *
 * Do not "simplify" this to a /AcroForm check. Verified against pdf-lib:
 * a PDF saved with object streams (its default) carries form fields while
 * the /AcroForm literal is absent from the raw bytes, so the one-literal
 * scan skips a document that genuinely needs renaming.
 */
function canHaveAcroForm(buf) {
  return buf.includes('/AcroForm') || buf.includes('/ObjStm')
}

async function prepPdfForDelivery(buf, filename) {
  if (!buf || buf.length < 5 || buf.slice(0, 5).toString() !== '%PDF-') {
    return buf
  }
  // PDFDocument.load is synchronous CPU work despite the async signature,
  // and it parses the whole file before we can ask whether it even has a
  // form. Skip that cost for PDFs that provably have none.
  if (!canHaveAcroForm(buf)) return buf
  try {
    const { PDFDocument, PDFName } = await import('pdf-lib')
    const doc = await PDFDocument.load(buf, { updateMetadata: false })
    const form = doc.getForm()
    const fields = form.getFields()
    if (fields.length === 0) return buf

    // Suffix derived from the filename — guaranteed unique per file
    // within the zip and human-readable if anyone inspects the form
    // tool's field list. Strip extension + replace anything outside
    // [a-zA-Z0-9_] so the resulting /T value is well-formed.
    const suffix = '__' + String(filename).replace(/\.[^.]+$/, '').replace(/[^a-zA-Z0-9_]/g, '_')

    let renamed = 0
    let lockedRO = 0
    for (const field of fields) {
      try {
        const partial = field.acroField.getPartialName() || field.getName() || 'field'
        field.acroField.setPartialName(partial + suffix)
        renamed++
      } catch (e) {
        // Some merged docs surface field shapes pdf-lib chokes on;
        // skip the offender rather than abort the whole file.
        console.warn(`prep: rename failed on a field in ${filename}: ${e.message}`)
      }
      try {
        field.enableReadOnly()
        lockedRO++
      } catch (e) {
        console.warn(`prep: enableReadOnly failed on a field in ${filename}: ${e.message}`)
      }
    }

    // updateFieldAppearances:false is critical here too — the default
    // save path regenerates /AP if a form was accessed via getForm(),
    // and that regeneration uses pdf-lib's font logic which would
    // clobber the PyMuPDF-rendered worker appearances.
    const out = Buffer.from(await doc.save({ updateFieldAppearances: false }))
    if (renamed === 0 && lockedRO === 0) {
      console.warn(`prep: no fields modified in ${filename} (renamed=0 lockedRO=0) — sending original`)
      return buf
    }
    return out
  } catch (e) {
    console.warn(`prep failed for ${filename}, sending original: ${e.message}`)
    return buf
  }
}

/**
 * Build the human-readable MANIFEST.txt content for a doc batch.
 * Plain text, monospace-friendly so it renders well in any reader
 * the prime contractor opens it with.
 */
function buildBatchManifest(filters, listing) {
  const lines = []
  const files   = listing.files   || []
  const missing = listing.missing || []
  const counts  = listing.counts  || {}

  const now = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/New_York',
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', hour12: false,
  }).format(new Date())

  lines.push('ONEIRO COLLECTION LLC — Document Delivery')
  lines.push('Generated: ' + now + ' ET')
  lines.push('Mode: ' + (filters.mode || 'unspecified'))
  const detail = []
  if (filters.contractors && filters.contractors.length) detail.push('contractors=[' + filters.contractors.join(', ') + ']')
  if (filters.doc_types   && filters.doc_types.length)   detail.push('doc_types=['   + filters.doc_types.join(', ')   + ']')
  if (filters.mode === 'date_range') detail.push(`dates=[${filters.date_start} → ${filters.date_end}]`)
  if (filters.mode === 'wo_numbers') detail.push('wos=[' + (filters.wo_ids || []).join(', ') + ']')
  if (detail.length) lines.push('Filters: ' + detail.join('  '))

  // Distinct WO count across the batch
  const allWos = new Set()
  files.forEach(f => (f.wo_ids || []).forEach(id => allWos.add(id)))
  lines.push(`Total: ${files.length} file${files.length === 1 ? '' : 's'} across ${allWos.size} work order${allWos.size === 1 ? '' : 's'}`)
  lines.push('')

  // Natural-sort comparator for WO numbers like "RM-43101", "PT-12345":
  // prefix alphabetical, then numeric portion ascending.
  const woCompare = (a, b) => {
    const aa = String(a || ''); const bb = String(b || '')
    const ap = aa.split('-')[0]; const bp = bb.split('-')[0]
    if (ap !== bp) return ap.localeCompare(bp)
    const an = parseInt(aa.split('-')[1] || '0', 10)
    const bn = parseInt(bb.split('-')[1] || '0', 10)
    if (an !== bn) return an - bn
    return aa.localeCompare(bb)
  }
  const dateCompare = (a, b) => String(a || '').localeCompare(String(b || ''))

  const sectionRule = '─'.repeat(40)
  const bullet      = '  - '

  // Each section: a labeled header line ("X for:") followed by a
  // bulleted list. Per-WO docs (CFR, Invoice) sort by WO ascending;
  // master-copy docs (Production Log, Sign-In, CP) sort by date.

  // ── Contractor Field Reports
  const cfrFiles = files.filter(f => f.doc_type === 'CFR')
    .sort((a, b) => woCompare((a.wo_ids || [])[0], (b.wo_ids || [])[0]))
  if (cfrFiles.length) {
    lines.push(sectionRule)
    lines.push(`Contractor Field Reports for: (${cfrFiles.length})`)
    lines.push(sectionRule)
    cfrFiles.forEach(f => {
      const woId = (f.wo_ids || [])[0] || '?'
      const loc  = f.location || ''
      const date = f.work_date || '?'
      const locPart = loc ? ` - ${loc}` : ''
      lines.push(`${bullet}${woId}${locPart} (${date})`)
    })
    lines.push('')
  }

  // ── Production Logs (crew chief shown so same-date multi-crew logs are
  // distinguishable — otherwise two identical dates look like a duplicate).
  const plFiles = files.filter(f => f.doc_type === 'Production Log')
    .sort((a, b) => dateCompare(a.work_date, b.work_date)
                    || String(a.crew_chief || '').localeCompare(String(b.crew_chief || '')))
  if (plFiles.length) {
    lines.push(sectionRule)
    lines.push(`Production Logs for: (${plFiles.length})`)
    lines.push(sectionRule)
    plFiles.forEach(f => {
      const chief = f.crew_chief
        ? `  —  Crew: ${String(f.crew_chief).replace(/([a-z0-9])([A-Z])/g, '$1 $2')}`
        : ''
      lines.push(`${bullet}${f.work_date || '?'}${chief}`)
    })
    lines.push('')
  }

  // ── Invoices
  const invFiles = files.filter(f => f.doc_type === 'Invoice')
    .sort((a, b) => woCompare((a.wo_ids || [])[0], (b.wo_ids || [])[0]))
  if (invFiles.length) {
    lines.push(sectionRule)
    lines.push(`Invoices for: (${invFiles.length})`)
    lines.push(sectionRule)
    invFiles.forEach(f => {
      const woId = (f.wo_ids || [])[0] || '?'
      const loc  = f.location || ''
      const locPart = loc ? ` - ${loc}` : ''
      lines.push(`${bullet}${woId}${locPart}`)
    })
    lines.push('')
  }

  // ── Certified Payroll
  const cpFiles = files.filter(f => f.doc_type === 'Certified Payroll')
    .sort((a, b) => dateCompare(a.work_date, b.work_date))
  if (cpFiles.length) {
    lines.push(sectionRule)
    lines.push(`Certified Payroll for Week(s) of: (${cpFiles.length})`)
    lines.push(sectionRule)
    cpFiles.forEach(f => {
      lines.push(`${bullet}${f.work_date || '?'}`)
    })
    lines.push('')
  }

  // ── Sign-In Sheets
  const siFiles = files.filter(f => f.doc_type === 'Sign-In')
    .sort((a, b) => dateCompare(a.work_date, b.work_date)
                 || String(a.contract_num || '').localeCompare(String(b.contract_num || ''))
                 || String(a.borough      || '').localeCompare(String(b.borough      || '')))
  if (siFiles.length) {
    lines.push(sectionRule)
    lines.push(`Sign in Sheets for: (${siFiles.length})`)
    lines.push(sectionRule)
    siFiles.forEach(f => {
      const date     = f.work_date    || '?'
      const contract = f.contract_num || '?'
      const borough  = f.borough      || '?'
      lines.push(`${bullet}${date} - ${contract}, ${borough}`)
    })
    lines.push('')
  }

  if (missing.length) {
    lines.push(sectionRule)
    lines.push(`NOT INCLUDED — requested but not yet generated/approved (${missing.length})`)
    lines.push(sectionRule)
    const sortedMissing = missing.slice().sort((a, b) =>
      woCompare(a.wo_id, b.wo_id) || String(a.doc_type).localeCompare(String(b.doc_type))
    )
    sortedMissing.forEach(m => {
      lines.push(`${bullet}${m.wo_id} — ${m.doc_type} — ${m.reason}`)
    })
    lines.push('')
  }

  if (listing.truncated) {
    lines.push('NOTE: Result truncated at the server limit. Narrow the filters and re-run for the rest.')
    lines.push('')
  }

  return lines.join('\n')
}


// ── QuickBooks Online integration routes ──────────────────────
//
// OAuth bootstrap is one-time per environment (or whenever the refresh
// token chain breaks). See docs/quickbooks_integration.md for the user
// setup flow.

// ── OAuth CSRF state helpers ──────────────────────────────────
//
// Intuit security requirement: the `state` parameter passed through
// the authorize → callback round-trip must be cryptographically random
// and verified server-side, so an attacker can't trick the user into
// completing an OAuth flow the user didn't initiate. We store the
// state in a short-lived HttpOnly cookie set when /api/qb/auth-start
// fires, and verify it matches the callback's `state` query param.
//
// 10-minute lifetime is long enough for any human OAuth flow and
// short enough that stale states age out quickly.
const QB_STATE_COOKIE = 'qb_oauth_state'
const QB_STATE_TTL_MS = 10 * 60 * 1000

function readCookie(req, name) {
  const header = req.headers.cookie || ''
  for (const part of header.split(';')) {
    const idx = part.indexOf('=')
    if (idx === -1) continue
    const k = part.slice(0, idx).trim()
    if (k === name) return decodeURIComponent(part.slice(idx + 1).trim())
  }
  return null
}

/**
 * GET /api/qb/auth-start
 * Kicks off the QBO OAuth 2.0 authorize-code flow. Admin clicks this
 * once when first connecting, and again any time auth needs to be
 * re-established (invalid_grant on refresh, etc.). Generates a CSRF
 * state token, parks it in an HttpOnly cookie, and includes the same
 * value in the authorize URL — verified back on /api/qb/auth-callback.
 */
app.get('/api/qb/auth-start', async (req, res) => {
  try {
    const state = crypto.randomBytes(24).toString('base64url')
    res.cookie(QB_STATE_COOKIE, state, {
      httpOnly: true,
      secure:   process.env.NODE_ENV === 'production',
      sameSite: 'lax',
      maxAge:   QB_STATE_TTL_MS,
      path:     '/api/qb',
    })
    const url = await buildAuthorizeUrl(state)
    res.redirect(url)
  } catch (err) {
    console.error('GET /api/qb/auth-start error:', err.message)
    res.status(500).send(`Failed to start QuickBooks authorization: ${err.message}`)
  }
})

/**
 * GET /api/qb/auth-callback?code=...&state=...&realmId=...
 * Intuit's redirect target after the user grants consent. Exchanges
 * the auth code for tokens and persists the (encrypted) refresh
 * token via Apps Script. Always responds with a 302 redirect (never
 * HTML containing the code) so the OAuth code can't leak to third
 * parties via Referer headers — Intuit security requirement.
 */
app.get('/api/qb/auth-callback', async (req, res) => {
  const { code, realmId, state, error: oauthError } = req.query
  // Always clear the state cookie on callback, success or failure —
  // single-use, prevents replay.
  res.clearCookie(QB_STATE_COOKIE, { path: '/api/qb' })

  // CSRF: verify the state echoed back matches the one we put in the
  // HttpOnly cookie. timingSafeEqual to avoid a timing oracle.
  const cookieState = readCookie(req, QB_STATE_COOKIE)
  const queryState  = String(state || '')
  if (!cookieState || !queryState || cookieState.length !== queryState.length ||
      !crypto.timingSafeEqual(Buffer.from(cookieState), Buffer.from(queryState))) {
    console.warn('[QB] auth-callback CSRF state mismatch — refusing')
    return res.redirect('/?qb=error&msg=' + encodeURIComponent('CSRF state mismatch — please start the connection again'))
  }
  if (oauthError) {
    return res.redirect(`/?qb=error&msg=${encodeURIComponent(String(oauthError))}`)
  }
  if (!code) {
    return res.redirect('/?qb=error&msg=' + encodeURIComponent('Missing authorization code'))
  }
  try {
    await exchangeAuthCode(String(code), String(realmId || ''))
    res.redirect('/?qb=connected')
  } catch (err) {
    console.error('GET /api/qb/auth-callback error:', err.message)
    res.redirect('/?qb=error&msg=' + encodeURIComponent(err.message))
  }
})

/**
 * GET /api/qb/status
 * Reports the QB connection state. React polls this to decide whether
 * to render the disconnected banner above the WO table.
 */
app.get('/api/qb/status', async (_req, res) => {
  try {
    const status = await checkQbConnection()
    res.json(status)
  } catch (err) {
    console.error('GET /api/qb/status error:', err.message)
    res.status(500).json({ connected: false, reason: 'error', error: err.message })
  }
})

/**
 * GET /api/qb/disconnect
 *
 * Disconnect webhook. Intuit calls this when a QBO admin revokes our
 * app from QBO → Apps → Connected Apps. Clears the stored refresh
 * token so future requests fail closed (showing the disconnected
 * banner) instead of trying to use a revoked grant.
 *
 * Intuit hits this as a plain GET with no signature, so anyone who
 * knows the URL can clear the token. The worst case is an admin has
 * to re-authorize — no data leak — so we accept the request as-is.
 * Validates the optional realmId query param against our configured
 * QB_REALM_ID when present, as a sanity check.
 */
app.get('/api/qb/disconnect', async (req, res) => {
  const { realmId } = req.query
  if (realmId && process.env.QB_REALM_ID && String(realmId) !== process.env.QB_REALM_ID) {
    console.warn(`[QB] disconnect: realmId ${realmId} doesn't match QB_REALM_ID — ignoring`)
    return res.status(400).send('Realm ID mismatch — refusing to clear token.')
  }
  try {
    await callAppsScript('set_qb_refresh_token', { token: '__cleared__' })
    console.log('[QB] disconnect: refresh token cleared')
    res.send(`
      <html><head><title>QuickBooks Disconnected</title>
      <style>body{font-family:system-ui;padding:40px;max-width:560px;margin:0 auto}
      h1{color:#0f172a}a{color:#1e40af}</style></head>
      <body>
        <h1>QuickBooks Disconnected</h1>
        <p>Your QuickBooks Online account has been disconnected from the Oneiro
           Operations Platform. To reconnect, an admin can visit
           <a href="/api/qb/auth-start">/api/qb/auth-start</a> from the platform.</p>
      </body></html>
    `)
  } catch (err) {
    console.error('GET /api/qb/disconnect error:', err.message)
    res.status(500).send(`Failed to clear QuickBooks token: ${err.message}`)
  }
})

/**
 * POST /api/qb/invoice/:woId
 * Generates a QB invoice for one Work Order. Idempotent — if the WO
 * already has an invoice number recorded, returns the existing one
 * without creating a duplicate.
 *
 * Response:
 *   { ok: true,  doc_number, qb_invoice_id, view_url, amount,
 *                already_invoiced: false }
 *   { ok: true,  doc_number, qb_invoice_id, view_url, amount,
 *                already_invoiced: true   }
 *   { ok: false, error, needs_pricing? }
 */
app.post('/api/qb/invoice/:woId', async (req, res) => {
  const woId = req.params.woId
  try {
    // 1) Fetch payload from Apps Script. Pre-flight check for
    //    "already invoiced" lives there — server-side guard against
    //    double billing even if the React side races.
    let payload = await callAppsScript('get_qb_invoice_payload', { wo_id: woId })

    if (payload.already_invoiced) {
      // Self-heal: an admin may have deleted the invoice in QB
      // without clearing the WO Tracker. Verify the recorded invoice
      // still exists; if not, clear the stale row and re-create.
      const stillExists = await qbInvoiceExists(payload.qb_invoice_id)
      if (stillExists) {
        return res.json({
          ok:               true,
          already_invoiced: true,
          doc_number:       payload.doc_number,
          qb_invoice_id:    payload.qb_invoice_id,
          view_url:         buildInvoiceViewUrl(payload.qb_invoice_id),
          amount:           payload.amount,
        })
      }
      console.warn(`[QB] WO ${woId} marked already-invoiced (#${payload.doc_number}, qb_id=${payload.qb_invoice_id}) but invoice is gone — clearing and re-generating`)
      await callAppsScript('clear_qb_invoice', { wo_id: woId })
      payload = await callAppsScript('get_qb_invoice_payload', { wo_id: woId })
      if (payload.already_invoiced) {
        // Defensive — shouldn't happen, but bail with a clear error
        // instead of silently looping.
        return res.status(500).json({
          ok: false,
          error: 'Failed to clear stale invoice record on WO Tracker — please clear cols 27–29 + 50 manually.',
        })
      }
    }

    if (Array.isArray(payload.needs_pricing) && payload.needs_pricing.length > 0) {
      return res.status(400).json({
        ok:    false,
        error: 'Some marking items can\'t be priced — fix them before invoicing.',
        needs_pricing: payload.needs_pricing,
      })
    }
    if (!payload.lines || payload.lines.length === 0) {
      return res.status(400).json({
        ok:    false,
        error: 'No priced marking items for this WO — nothing to invoice.',
      })
    }

    // 2) Resolve customer (cache → query → error if missing)
    const customer = await findCustomerByName(payload.contractor)

    // 3) Create the invoice
    const result = await createInvoiceForWO(payload, customer.id)

    // 4) Record the result back on the WO Tracker
    await callAppsScript('record_qb_invoice', {
      wo_id:         woId,
      doc_number:    result.doc_number,
      qb_invoice_id: result.qb_invoice_id,
      amount:        payload.totals.revenue,
    })

    res.json({
      ok:               true,
      already_invoiced: false,
      doc_number:       result.doc_number,
      qb_invoice_id:    result.qb_invoice_id,
      view_url:         result.view_url,
      amount:           payload.totals.revenue,
    })
  } catch (err) {
    console.error(`POST /api/qb/invoice/${woId} error:`, err.message)
    const status = err.message?.includes('QB_NOT_CONNECTED') ? 401 : 500
    res.status(status).json({ ok: false, error: err.message })
  }
})


// ── Access gate ───────────────────────────────────────────────
// Lightweight shared-code access control — a door code, not real auth
// (no accounts, no passwords). Two codes live in Railway env vars:
// ACCESS_KEY_ADMIN unlocks the full app, ACCESS_KEY_CREW unlocks the
// crew view (Nav + Field Report + Sign-In only). A visitor arrives with
// a code in their link (`/?key=<code>`); the client redeems it at
// /api/access/login, which validates against the env vars and, on a
// match, drops a long-lived HttpOnly cookie holding the raw code. Every
// page load then calls /api/access/session, which RE-checks that cookie
// against the CURRENT env vars — so rotating a code in Railway instantly
// locks out everyone holding the old link (e.g. a departed employee),
// with no redeploy.
//
// Fail-open by design: if NEITHER code is configured the gate is off and
// everyone is treated as admin. That keeps local dev open and means
// deploying this code can't brick the site before the env vars are set —
// the gate only switches on once you set a code in Railway. The raw
// /api/* data routes are intentionally NOT gated here (automations like
// the QuickBooks callback and Apps Script uploads call them server-to-
// server); this pass gates the UI only.
const ACCESS_COOKIE = 'oneiro_access'
const ACCESS_TTL_MS = 180 * 24 * 60 * 60 * 1000  // 180 days

function safeEqual(a, b) {
  const ba = Buffer.from(String(a))
  const bb = Buffer.from(String(b))
  if (ba.length !== bb.length) return false
  return crypto.timingSafeEqual(ba, bb)
}

// Presented code → role ('admin' | 'crew' | null), compared in constant
// time. Blank env vars never match, so an unset code can't be satisfied
// by an empty cookie. Both unset → gate disabled (everyone is admin).
function roleForCode(code) {
  const admin = process.env.ACCESS_KEY_ADMIN || ''
  const crew  = process.env.ACCESS_KEY_CREW  || ''
  if (!admin && !crew) return 'admin'
  const c = String(code || '')
  if (admin && safeEqual(c, admin)) return 'admin'
  if (crew  && safeEqual(c, crew))  return 'crew'
  return null
}

// POST /api/access/login { key } — validate a presented code; on success
// set the access cookie and return the granted role, else 401 + no cookie.
app.post('/api/access/login', (req, res) => {
  const key  = req.body?.key
  const role = roleForCode(key)
  if (!role) return res.status(401).json({ role: null })
  res.cookie(ACCESS_COOKIE, String(key ?? ''), {
    httpOnly: true,
    secure:   process.env.NODE_ENV === 'production',
    sameSite: 'lax',
    maxAge:   ACCESS_TTL_MS,
    path:     '/',
  })
  res.json({ role })
})

// GET /api/access/session — re-validate the cookie against the CURRENT
// env vars on every call, so a rotated code takes effect immediately.
app.get('/api/access/session', (req, res) => {
  res.json({ role: roleForCode(readCookie(req, ACCESS_COOKIE)) })
})

// POST /api/access/logout — clear the cookie (sign out / switch links).
app.post('/api/access/logout', (req, res) => {
  res.clearCookie(ACCESS_COOKIE, { path: '/' })
  res.json({ role: null })
})


// ── Static file serving (production only) ────────────────────
if (process.env.NODE_ENV === 'production') {
  const distDir = path.join(__dirname, 'dist')

  // Self-hosted scanner assets (OpenCV.js ~9 MB + jscanify). Cache hard
  // so the crew downloads them once per device, not every scan — these
  // are pinned versions, so bump the filename if the version changes.
  app.use('/vendor', express.static(path.join(distDir, 'vendor'), {
    immutable: true,
    maxAge:    '365d',
  }))

  app.use(express.static(distDir))

  // All non-API routes → serve index.html (client-side routing)
  app.get('*', (req, res) => {
    if (!req.path.startsWith('/api')) {
      res.sendFile(path.join(distDir, 'index.html'))
    }
  })
}


// ── Start ─────────────────────────────────────────────────────
// QB Item IDs sanity-check. Logs loudly when QB env vars are present
// but qbItems.js still has placeholders. Non-fatal: the rest of the
// app boots normally and only the QB invoice endpoints will fail when
// called (with the same clear error message). This lets us partially
// configure QB without taking the whole dashboard down.
try {
  assertQbItemsConfigured()
} catch (err) {
  console.error('⚠ QB items misconfigured (QB invoicing disabled):', err.message)
}
const qbCfg = qbConfigStatus()
console.log(qbCfg.configured
  ? `   QB:   configured (${qbCfg.sandbox ? 'sandbox' : 'production'})`
  : `   QB:   disabled (missing env vars: ${qbCfg.missing.join(', ')})`)

app.listen(PORT, () => {
  console.log(`🚀 Oneiro Ops web server running on port ${PORT}`)
  console.log(`   Mode: ${process.env.NODE_ENV || 'development'}`)
  if (process.env.NODE_ENV !== 'production') {
    console.log(`   API:  http://localhost:${PORT}/api/health`)
    console.log(`   App:  http://localhost:5173  (Vite dev server)`)
  }
})
