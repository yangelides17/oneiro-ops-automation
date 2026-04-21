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

  // Apps Script always returns 200; errors come in the JSON body
  const json = await res.json()
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
