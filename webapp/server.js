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
