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

// Hardcoded fallbacks — used only if the OIDC discovery doc fetch
// fails. The live values come from getOAuthEndpoints() below, which
// consults Intuit's OpenID Connect discovery document so we always
// use whatever endpoint URLs Intuit currently advertises.
const FALLBACK_TOKEN_URL     = 'https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer'
const FALLBACK_AUTHORIZE_URL = 'https://appcenter.intuit.com/connect/oauth2'
const DISCOVERY_URL          = 'https://developer.api.intuit.com/.well-known/openid_configuration/'
const SCOPE                  = 'com.intuit.quickbooks.accounting'

let _discoveryCache = null   // { authorization_endpoint, token_endpoint, revocation_endpoint }

async function getOAuthEndpoints() {
  if (_discoveryCache) return _discoveryCache
  try {
    const res = await fetch(DISCOVERY_URL, { headers: { 'Accept': 'application/json' } })
    if (res.ok) {
      const json = await res.json()
      _discoveryCache = {
        authorization_endpoint: json.authorization_endpoint || FALLBACK_AUTHORIZE_URL,
        token_endpoint:         json.token_endpoint         || FALLBACK_TOKEN_URL,
        revocation_endpoint:    json.revocation_endpoint    || null,
      }
      console.log('[QB] OIDC discovery loaded — auth=', _discoveryCache.authorization_endpoint,
                  ' token=', _discoveryCache.token_endpoint)
      return _discoveryCache
    }
    console.warn('[QB] discovery fetch returned non-OK:', res.status)
  } catch (err) {
    console.warn('[QB] discovery fetch failed:', err.message)
  }
  _discoveryCache = {
    authorization_endpoint: FALLBACK_AUTHORIZE_URL,
    token_endpoint:         FALLBACK_TOKEN_URL,
    revocation_endpoint:    null,
  }
  return _discoveryCache
}

export function qbConfigStatus() {
  const missing = []
  if (!CLIENT_ID)     missing.push('QB_CLIENT_ID')
  if (!CLIENT_SECRET) missing.push('QB_CLIENT_SECRET')
  if (!REDIRECT_URI)  missing.push('QB_REDIRECT_URI')
  if (!REALM_ID)      missing.push('QB_REALM_ID')
  return { configured: missing.length === 0, missing, sandbox: BASE_URL.includes('sandbox') }
}

// ── Refresh-token encryption (AES-256-GCM) ────────────────────────
// Intuit security requirement: store the OAuth refresh token encrypted
// with a symmetric algorithm. We use AES-256-GCM with a per-encrypt
// random 96-bit IV. The encrypted blob is what gets written to Apps
// Script Properties — Apps Script never sees plaintext. Key lives in
// the QB_TOKEN_ENCRYPTION_KEY env var (32 raw bytes, base64-encoded).
//
// Generate a key once with:
//   node -e "console.log(require('crypto').randomBytes(32).toString('base64'))"
// and add the output to webapp/.env and Railway env. Rotating the key
// invalidates any stored token and forces a re-auth — fine, cheap.
//
// Storage layout (base64):
//   bytes 0-11  : 96-bit IV
//   bytes 12-27 : 128-bit auth tag
//   bytes 28+   : ciphertext

const ENC_PREFIX = 'enc:v1:'   // distinguishes encrypted blob from any
                               // legacy plaintext that might exist

function _getEncKey() {
  const raw = process.env.QB_TOKEN_ENCRYPTION_KEY
  if (!raw) {
    throw new Error(
      'QB_TOKEN_ENCRYPTION_KEY env var not set. Generate one with: ' +
      'node -e "console.log(require(\'crypto\').randomBytes(32).toString(\'base64\'))"'
    )
  }
  const key = Buffer.from(raw, 'base64')
  if (key.length !== 32) {
    throw new Error(`QB_TOKEN_ENCRYPTION_KEY must decode to 32 bytes (got ${key.length})`)
  }
  return key
}

function encryptToken(plaintext) {
  const key = _getEncKey()
  const iv  = crypto.randomBytes(12)
  const cipher = crypto.createCipheriv('aes-256-gcm', key, iv)
  const ct = Buffer.concat([cipher.update(String(plaintext), 'utf8'), cipher.final()])
  const tag = cipher.getAuthTag()
  return ENC_PREFIX + Buffer.concat([iv, tag, ct]).toString('base64')
}

function decryptToken(stored) {
  if (!stored) return ''
  // Tolerate any legacy plaintext (shouldn't exist since this lands
  // before first auth, but defensive — and we explicitly mark our
  // encrypted blobs with ENC_PREFIX so future versions can migrate).
  if (typeof stored !== 'string' || !stored.startsWith(ENC_PREFIX)) {
    return stored
  }
  const key = _getEncKey()
  const data = Buffer.from(stored.slice(ENC_PREFIX.length), 'base64')
  const iv  = data.subarray(0, 12)
  const tag = data.subarray(12, 28)
  const ct  = data.subarray(28)
  const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv)
  decipher.setAuthTag(tag)
  return Buffer.concat([decipher.update(ct), decipher.final()]).toString('utf8')
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
export async function buildAuthorizeUrl(state) {
  const cfg = qbConfigStatus()
  if (!cfg.configured) throw new Error('QB env vars missing: ' + cfg.missing.join(', '))
  const endpoints = await getOAuthEndpoints()
  const params = new URLSearchParams({
    client_id:     CLIENT_ID,
    response_type: 'code',
    scope:         SCOPE,
    redirect_uri:  REDIRECT_URI,
    state:         state,
  })
  return `${endpoints.authorization_endpoint}?${params.toString()}`
}

// ── OAuth: exchange authorization code for tokens (one-time per env) ─
export async function exchangeAuthCode(code, realmId) {
  if (realmId && realmId !== REALM_ID) {
    throw new Error(
      `OAuth callback realmId=${realmId} doesn't match QB_REALM_ID=${REALM_ID}. ` +
      `Authorize against the correct QBO company or fix the env var.`
    )
  }
  const endpoints = await getOAuthEndpoints()
  const body = new URLSearchParams({
    grant_type:   'authorization_code',
    code,
    redirect_uri: REDIRECT_URI,
  })
  const res = await fetch(endpoints.token_endpoint, {
    method:  'POST',
    headers: {
      'Authorization': 'Basic ' + Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString('base64'),
      'Content-Type':  'application/x-www-form-urlencoded',
      'Accept':        'application/json',
    },
    body: body.toString(),
  })
  const exchangeTid = res.headers.get('intuit_tid') || ''
  const json = await res.json()
  if (!res.ok) {
    console.error(`[QB] auth-code exchange failed status=${res.status} intuit_tid=${exchangeTid}`)
    throw new Error(`QB token exchange failed (${res.status}) [intuit_tid=${exchangeTid}]: ${JSON.stringify(json)}`)
  }
  if (!json.refresh_token) throw new Error('QB token exchange did not return refresh_token')

  await callAppsScript('set_qb_refresh_token', { token: encryptToken(json.refresh_token) })
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
    const stored = String(cur.token || '').trim()
    if (!stored || stored === '__cleared__') throw new Error('QB_NOT_CONNECTED')
    let oldToken
    try {
      oldToken = decryptToken(stored)
    } catch (err) {
      console.log(`[QB] refresh failed to decrypt stored token: ${err.message}`)
      throw new Error('QB_NOT_CONNECTED')
    }
    console.log(`[QB] refresh start, old_hash=${_hashToken(oldToken)}`)

    const endpoints = await getOAuthEndpoints()
    const body = new URLSearchParams({
      grant_type:    'refresh_token',
      refresh_token: oldToken,
    })
    const res = await fetch(endpoints.token_endpoint, {
      method:  'POST',
      headers: {
        'Authorization': 'Basic ' + Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString('base64'),
        'Content-Type':  'application/x-www-form-urlencoded',
        'Accept':        'application/json',
      },
      body: body.toString(),
    })
    const tokenTid = res.headers.get('intuit_tid') || ''
    const json = await res.json()
    if (!res.ok) {
      console.log(`[QB] refresh failed: ${res.status} intuit_tid=${tokenTid} ${JSON.stringify(json)}`)
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
    await callAppsScript('set_qb_refresh_token', { token: encryptToken(newRefresh) })

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
  // intuit_tid (Intuit Transaction ID) is in every QB API response
  // header. Capture for log correlation + include in thrown errors so
  // admins can hand it to Intuit support for fast root-cause lookup.
  const intuitTid = res.headers.get('intuit_tid') || ''
  const text = await res.text()
  let json = null
  try { json = text ? JSON.parse(text) : null } catch (_) { /* leave null */ }
  if (!res.ok) {
    const fault = json && json.Fault ? JSON.stringify(json.Fault) : text.slice(0, 300)
    console.error(`[QB] ${method} ${path} failed status=${res.status} intuit_tid=${intuitTid} fault=${fault}`)
    throw new Error(`QB ${method} ${path} failed (${res.status}) [intuit_tid=${intuitTid}]: ${fault}`)
  }
  console.log(`[QB] ${method} ${path} ok status=${res.status} intuit_tid=${intuitTid}`)
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

// ── Sales Term lookup (Net 30) ────────────────────────────────────
// Setting the term explicitly on the invoice payload is more reliable
// than relying on QB to inherit it from the customer's default — when
// EITHER SalesTermRef or DueDate is omitted, certain QBO configurations
// just leave the term blank on the rendered invoice. We resolve "Net 30"
// once per webapp process, cache the ID, and attach it to every
// generated invoice. If the company has renamed/deleted the term, the
// invoice ships without one (and QB might apply customer default).
let _net30TermId    = null
let _net30Resolved  = false

async function getNet30TermId() {
  if (_net30Resolved) return _net30TermId
  try {
    const json = await qbFetch('GET', '/query', {
      query: { query: "SELECT Id, Name FROM Term WHERE Name = 'Net 30'" },
    })
    const term = (json && json.QueryResponse && json.QueryResponse.Term || [])[0]
    if (term && term.Id) {
      _net30TermId = String(term.Id)
      console.log(`[QB] resolved Net 30 SalesTerm: id=${_net30TermId}`)
    } else {
      console.warn('[QB] no SalesTerm named "Net 30" found — invoices will fall back to whatever QB applies')
    }
  } catch (err) {
    console.warn(`[QB] Net 30 lookup failed: ${err.message}`)
  }
  _net30Resolved = true
  return _net30TermId
}

// ── Invoice creation ──────────────────────────────────────────────
async function buildQbInvoice(payload, customerId) {
  const itemLines = payload.lines.map(l => {
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

  // ── WO info header (Path B) ─────────────────────────────────
  // Single DescriptionOnly line at the top of the items table. The
  // Description carries 5 fields separated by newlines — renders as
  // a single tall row above the priced lines. Sits visually inside
  // the line-items table so it's universally placed across every QB
  // template, with no other column data attached.
  const headerText = [
    payload.wo_id        && `Work Order: ${payload.wo_id}`,
    payload.contract_num && `Contract #: ${payload.contract_num}`,
    payload.borough      && `Borough: ${payload.borough}`,
    payload.location     && `Location: ${payload.location}`,
    payload.work_start && payload.work_end &&
      `Work Period: ${payload.work_start} – ${payload.work_end}`,
  ].filter(Boolean).join('\n')
  const headerRow = headerText
    ? [{ Amount: 0, DetailType: 'DescriptionOnly', Description: headerText }]
    : []

  // Explicitly attach the Net 30 SalesTerm so it renders on the
  // invoice. We still leave DueDate unset — QB computes it from the
  // term. If the lookup didn't find a "Net 30" term in this QB
  // company, we omit SalesTermRef and let QB apply whatever default
  // it can (typically the customer's default Term).
  const termId = await getNet30TermId()
  // Invoice date = the date this invoice is generated (NYC time), not
  // the WO's work_end date. Crews often complete a WO days before
  // billing goes out, and AR aging starts from the invoice date.
  // The Work Period is still visible in the header description row.
  const txnDate = new Date().toLocaleDateString('en-CA', { timeZone: 'America/New_York' })
  return {
    CustomerRef: { value: String(customerId) },
    TxnDate:     txnDate,
    PrivateNote: `WO ${payload.wo_id} · ${payload.location}`,
    ...(termId ? { SalesTermRef: { value: termId } } : {}),
    Line: [...headerRow, ...itemLines],
  }
}

export async function createInvoiceForWO(payload, customerId) {
  // We intentionally do NOT pass a RequestId here. QB's idempotency
  // record persists past invoice deletion — using a deterministic
  // RequestId (e.g. sha256 of WO + work_end + total) means once we've
  // created an invoice for a given WO, QB will keep handing back the
  // SAME (potentially deleted) invoice ID forever. The React UI
  // disables the Generate button during the in-flight request and the
  // server doesn't auto-retry non-401 errors, so RequestId wasn't
  // protecting us from anything that wasn't already covered.
  const invoice = await buildQbInvoice(payload, customerId)
  const result  = await qbFetch('POST', '/invoice', { body: invoice })
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
    const stored = String(refresh && refresh.token || '').trim()
    if (!stored || stored === '__cleared__') {
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

// ── Verify a stored QB Invoice ID still exists ────────────────────
// Used by the "already invoiced" auto-heal path: if an admin deletes
// the invoice in QB without clearing the WO Tracker, this returns
// false and the caller clears the sheet + creates a fresh one. Any
// fetch error other than a clear "not found" is treated as "still
// exists" — we'd rather skip an unnecessary re-create on a transient
// glitch than create a duplicate.
export async function qbInvoiceExists(qbInvoiceId) {
  if (!qbInvoiceId) return false
  try {
    const json = await qbFetch('GET', `/invoice/${encodeURIComponent(qbInvoiceId)}`)
    return !!(json && json.Invoice && json.Invoice.Id)
  } catch (err) {
    const msg = String(err && err.message || err)
    // QB returns code 610 "Object Not Found" with HTTP 400 when an
    // entity is gone. Some shapes also include "ObjectNotFound" or
    // the literal substring "Object Not Found" in the Fault detail.
    if (msg.includes('"code":"610"') ||
        msg.includes('Object Not Found') ||
        msg.includes('ObjectNotFound')) {
      return false
    }
    // Anything else — log and assume still exists, so we don't dupe.
    console.warn(`[QB] qbInvoiceExists ${qbInvoiceId} ambiguous: ${msg}`)
    return true
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
