/**
 * QuickBooks Online (QBO) API client.
 *
 * Responsibilities:
 *   - OAuth 2.0 authorize-code exchange + refresh token rotation
 *   - Persisting the rotated refresh token via Apps Script
 *   - Generic qbFetch wrapper with auto-retry on 401
 *   - Customer-by-name lookup with cache
 *   - Invoice creation from the Apps Script aggregator payload
 *
 * Token-storage architecture:
 *   - access_token:    in-memory only (60-min TTL; cheap to refresh)
 *   - refresh_token:   Apps Script PropertiesService (durable across
 *                      restarts + writable from running webapp)
 *
 * Persistence ordering (critical for the rotation chain):
 *   1. POST oauth2/v1/tokens/bearer (grant_type=refresh_token)
 *   2. Receive { access_token, refresh_token: new, expires_in }
 *   3. Persist new refresh_token to Apps Script               ← must succeed
 *   4. Update in-memory access_token + expiry
 *   5. Return the access_token to the caller
 *
 * If step 3 fails we abort BEFORE step 4 — the old refresh token on
 * Intuit's side stays valid until you use the new one, so the next
 * attempt retries cleanly from step 1.
 *
 * See docs/quickbooks_integration.md for the user-side QB setup.
 */

import crypto from 'crypto'
import { QB_ITEMS } from './qbItems.js'

// ── Config from env ───────────────────────────────────────────────
const CLIENT_ID     = process.env.QB_CLIENT_ID     || ''
const CLIENT_SECRET = process.env.QB_CLIENT_SECRET || ''
const REDIRECT_URI  = process.env.QB_REDIRECT_URI  || ''
const REALM_ID      = process.env.QB_REALM_ID      || ''
const BASE_URL      = process.env.QB_BASE_URL      || 'https://quickbooks.api.intuit.com'
// The QB UI base differs between sandbox and production — used to build
// view URLs for invoices that admins can click to open in QBO.
const QB_APP_BASE   = BASE_URL.includes('sandbox')
  ? 'https://app.sandbox.qbo.intuit.com'
  : 'https://app.qbo.intuit.com'

const TOKEN_URL     = 'https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer'
const AUTHORIZE_URL = 'https://appcenter.intuit.com/connect/oauth2'
const SCOPE         = 'com.intuit.quickbooks.accounting'

export function qbConfigStatus() {
  const missing = []
  if (!CLIENT_ID)     missing.push('QB_CLIENT_ID')
  if (!CLIENT_SECRET) missing.push('QB_CLIENT_SECRET')
  if (!REDIRECT_URI)  missing.push('QB_REDIRECT_URI')
  if (!REALM_ID)      missing.push('QB_REALM_ID')
  return { configured: missing.length === 0, missing, sandbox: BASE_URL.includes('sandbox') }
}

// ── Apps Script proxy (token + customer cache live there) ─────────
async function callAppsScript(action, data = null) {
  const url = process.env.APPS_SCRIPT_URL
  const key = process.env.APPS_SCRIPT_KEY
  if (!url || !key) throw new Error('APPS_SCRIPT_URL or APPS_SCRIPT_KEY env var not set')
  const body = { action, key, ...(data ? { data } : {}) }
  const res = await fetch(url, {
    method:  'POST',
    headers: { 'Content-Type': 'application/json' },
    body:    JSON.stringify(body),
  })
  const text = await res.text()
  let json
  try { json = JSON.parse(text) }
  catch (_) { throw new Error(`Apps Script returned non-JSON for ${action}: ${text.slice(0, 200)}`) }
  if (json.error) throw new Error(json.error)
  return json
}

// ── In-memory access token state ──────────────────────────────────
let _accessToken    = ''
let _accessExpiry   = 0     // ms epoch
let _refreshing     = null  // promise mutex — prevents concurrent refreshes

function _hashToken(t) {
  return t ? crypto.createHash('sha256').update(t).digest('hex').slice(0, 8) : 'empty'
}

// ── OAuth: authorize URL ──────────────────────────────────────────
export function buildAuthorizeUrl(state) {
  const cfg = qbConfigStatus()
  if (!cfg.configured) throw new Error('QB env vars missing: ' + cfg.missing.join(', '))
  const params = new URLSearchParams({
    client_id:     CLIENT_ID,
    response_type: 'code',
    scope:         SCOPE,
    redirect_uri:  REDIRECT_URI,
    state:         state,
  })
  return `${AUTHORIZE_URL}?${params.toString()}`
}

// ── OAuth: exchange authorization code for tokens (one-time per env) ─
export async function exchangeAuthCode(code, realmId) {
  if (realmId && realmId !== REALM_ID) {
    throw new Error(
      `OAuth callback realmId=${realmId} doesn't match QB_REALM_ID=${REALM_ID}. ` +
      `Authorize against the correct QBO company or fix the env var.`
    )
  }
  const body = new URLSearchParams({
    grant_type:   'authorization_code',
    code,
    redirect_uri: REDIRECT_URI,
  })
  const res = await fetch(TOKEN_URL, {
    method:  'POST',
    headers: {
      'Authorization': 'Basic ' + Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString('base64'),
      'Content-Type':  'application/x-www-form-urlencoded',
      'Accept':        'application/json',
    },
    body: body.toString(),
  })
  const json = await res.json()
  if (!res.ok) throw new Error(`QB token exchange failed (${res.status}): ${JSON.stringify(json)}`)
  if (!json.refresh_token) throw new Error('QB token exchange did not return refresh_token')

  await callAppsScript('set_qb_refresh_token', { token: json.refresh_token })
  _accessToken  = json.access_token
  _accessExpiry = Date.now() + (json.expires_in - 60) * 1000  // 60s safety margin
  console.log(`[QB] connect ok, refresh_hash=${_hashToken(json.refresh_token)}`)
  return { ok: true }
}

// ── OAuth: refresh access token (called lazily before any QB call) ─
async function refreshAccessToken() {
  // De-dupe concurrent refreshes within this process
  if (_refreshing) return _refreshing

  _refreshing = (async () => {
    const cur = await callAppsScript('get_qb_refresh_token')
    const oldToken = String(cur.token || '').trim()
    if (!oldToken) throw new Error('QB_NOT_CONNECTED')
    console.log(`[QB] refresh start, old_hash=${_hashToken(oldToken)}`)

    const body = new URLSearchParams({
      grant_type:    'refresh_token',
      refresh_token: oldToken,
    })
    const res = await fetch(TOKEN_URL, {
      method:  'POST',
      headers: {
        'Authorization': 'Basic ' + Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString('base64'),
        'Content-Type':  'application/x-www-form-urlencoded',
        'Accept':        'application/json',
      },
      body: body.toString(),
    })
    const json = await res.json()
    if (!res.ok) {
      console.log(`[QB] refresh failed: ${res.status} ${JSON.stringify(json)}`)
      // invalid_grant → admin must re-authorize
      throw new Error('QB_NOT_CONNECTED')
    }
    if (!json.access_token) throw new Error('QB refresh did not return access_token')

    const newRefresh = json.refresh_token || oldToken  // QB always rotates, but be defensive

    // ── Persist FIRST, then update in-memory state ─────────────
    // If this fails, the in-memory access token isn't updated — the
    // next call retries from scratch using the still-valid old
    // refresh token (Intuit only invalidates the old once you actually
    // use the new one). No silent invalidation possible.
    await callAppsScript('set_qb_refresh_token', { token: newRefresh })

    _accessToken  = json.access_token
    _accessExpiry = Date.now() + (json.expires_in - 60) * 1000
    console.log(`[QB] refresh ok, new_hash=${_hashToken(newRefresh)}, persisted=true, expires_in=${json.expires_in}`)
    return _accessToken
  })()

  try {
    return await _refreshing
  } finally {
    _refreshing = null
  }
}

async function getAccessToken() {
  if (_accessToken && Date.now() < _accessExpiry) return _accessToken
  return refreshAccessToken()
}

// ── Generic QB API request wrapper ────────────────────────────────
async function qbFetch(method, path, { body, query, requestId } = {}) {
  const cfg = qbConfigStatus()
  if (!cfg.configured) throw new Error('QB_NOT_CONFIGURED: ' + cfg.missing.join(', '))

  const qs = new URLSearchParams({ minorversion: '75', ...(query || {}) })
  if (requestId) qs.set('requestid', requestId)
  const url = `${BASE_URL}/v3/company/${REALM_ID}${path}?${qs.toString()}`

  const doFetch = async (token) => {
    const res = await fetch(url, {
      method,
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type':  'application/json',
        'Accept':        'application/json',
      },
      body: body ? JSON.stringify(body) : undefined,
    })
    return res
  }

  let token = await getAccessToken()
  let res = await doFetch(token)
  if (res.status === 401) {
    // Token may have been revoked; force a refresh and retry once
    _accessToken = ''
    token = await refreshAccessToken()
    res = await doFetch(token)
  }
  const text = await res.text()
  let json = null
  try { json = text ? JSON.parse(text) : null } catch (_) { /* leave null */ }
  if (!res.ok) {
    const fault = json && json.Fault ? JSON.stringify(json.Fault) : text.slice(0, 300)
    throw new Error(`QB ${method} ${path} failed (${res.status}): ${fault}`)
  }
  return json
}

// ── Customer lookup with cache ────────────────────────────────────
export async function findCustomerByName(name) {
  const normalized = String(name || '').trim()
  if (!normalized) throw new Error('findCustomerByName: name required')

  // 1) Cache hit?
  const cached = await callAppsScript('get_qb_customer_id', { name: normalized })
  if (cached && cached.id) return { id: String(cached.id), name: normalized }

  // 2) Query QB. Single-quotes inside the SQL-ish query must be escaped.
  const safeName = normalized.replace(/'/g, "\\'")
  const q = `SELECT Id, DisplayName FROM Customer WHERE DisplayName = '${safeName}'`
  const json = await qbFetch('GET', '/query', { query: { query: q } })
  const rows = (json && json.QueryResponse && json.QueryResponse.Customer) || []
  if (rows.length === 0) {
    throw new Error(
      `Customer "${normalized}" not found in QuickBooks. ` +
      `Add it via the QBO UI (Sales → Customers → New) first.`
    )
  }
  const customer = rows[0]
  const id = String(customer.Id)

  // 3) Cache for next time
  await callAppsScript('set_qb_customer_id', { name: normalized, id })
  return { id, name: normalized }
}

// ── Invoice creation ──────────────────────────────────────────────
function addDays(iso, n) {
  const d = new Date(iso + 'T12:00:00')
  d.setDate(d.getDate() + n)
  return d.toISOString().slice(0, 10)
}

function buildQbInvoice(payload, customerId) {
  const lines = payload.lines.map(l => {
    const itemId = QB_ITEMS[l.group]
    if (!itemId) throw new Error(`QB_ITEMS["${l.group}"] not configured in qbItems.js`)
    return {
      Amount:      l.amount,
      DetailType:  'SalesItemLineDetail',
      Description: l.description,
      SalesItemLineDetail: {
        ItemRef:   { value: String(itemId) },
        Qty:       l.qty,
        UnitPrice: l.rate,
      },
    }
  })
  return {
    CustomerRef:  { value: String(customerId) },
    TxnDate:      payload.work_end,
    DueDate:      addDays(payload.work_end, 30),
    PrivateNote:  `WO ${payload.wo_id} · ${payload.location}`,
    CustomerMemo: { value: `Work Order ${payload.wo_id}` },
    Line:         lines,
  }
}

export async function createInvoiceForWO(payload, customerId) {
  // Deterministic RequestId — same WO + work-end + total dedupes at QB
  // for accidental retries or double-clicks.
  const requestId = crypto.createHash('sha256')
    .update(`${payload.wo_id}|${payload.work_end}|${payload.totals.revenue}`)
    .digest('hex').slice(0, 36)

  const invoice = buildQbInvoice(payload, customerId)
  const result  = await qbFetch('POST', '/invoice', { body: invoice, requestId })
  const created = result && result.Invoice
  if (!created || !created.Id) {
    throw new Error('QB invoice POST returned no Invoice.Id: ' + JSON.stringify(result).slice(0, 300))
  }

  const qbId      = String(created.Id)
  const docNumber = String(created.DocNumber || created.Id)
  const viewUrl   = `${QB_APP_BASE}/app/invoice?txnId=${qbId}`
  return { qb_invoice_id: qbId, doc_number: docNumber, view_url: viewUrl, raw: created }
}

// ── Connection health check ───────────────────────────────────────
export async function checkQbConnection() {
  const cfg = qbConfigStatus()
  if (!cfg.configured) {
    return { connected: false, reason: 'env_missing', missing: cfg.missing, sandbox: cfg.sandbox }
  }
  try {
    const refresh = await callAppsScript('get_qb_refresh_token')
    if (!refresh || !refresh.token) {
      return { connected: false, reason: 'not_authorized', sandbox: cfg.sandbox }
    }
    // Cheap call: fetch CompanyInfo. Will trigger a refresh if needed.
    await qbFetch('GET', `/companyinfo/${REALM_ID}`)
    return { connected: true, sandbox: cfg.sandbox }
  } catch (err) {
    const msg = err.message || String(err)
    if (msg.includes('QB_NOT_CONNECTED')) {
      return { connected: false, reason: 'not_authorized', sandbox: cfg.sandbox }
    }
    return { connected: false, reason: 'error', error: msg, sandbox: cfg.sandbox }
  }
}

// ── Build a QB invoice view URL from a stored qb_invoice_id ───────
// Used by /api/dashboard to decorate WO rows that already have an
// invoice. Pure function — no env or network. Safe to call when QB
// isn't configured.
export function buildInvoiceViewUrl(qbInvoiceId) {
  if (!qbInvoiceId) return null
  return `${QB_APP_BASE}/app/invoice?txnId=${qbInvoiceId}`
}
