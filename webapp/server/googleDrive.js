/**
 * Direct Google Drive access for the web service.
 *
 * Why: preview bytes currently travel Drive → Apps Script → base64 (+33%)
 * → JSON → Node. That path pays a ~1.2-2.4s script-load tax, cannot stream,
 * and dies past Apps Script's ~30-35s response ceiling — measured at a 21%
 * failure rate, the worst in the app. Reading Drive directly removes all
 * four problems at once.
 *
 * Auth: the same service account the Python worker already uses
 * (GOOGLE_SERVICE_ACCOUNT_JSON). That account demonstrably reads these exact
 * folders in production — workers/watch_and_fill.py:44 watches the same four
 * subfolders that Code.js:4552 lists approvals from, and downloads from them
 * via files().get_media().
 *
 * Scope is deliberately drive.readonly, narrower than the worker's full
 * `drive`. Previews never write. Note this is LESS authority than the current
 * arrangement, not more: Apps Script runs as USER_DEPLOYING — the owner's own
 * account, which can see their entire Drive — whereas a service account sees
 * only what has been shared with it.
 *
 * No SDK dependency on purpose. A service-account grant is an RS256-signed
 * JWT exchanged for an access token; Node's crypto does RS256 natively, so
 * this is ~40 lines rather than a package in a credential path.
 */
import crypto from 'crypto'

const TOKEN_URL = 'https://oauth2.googleapis.com/token'
const SCOPE     = 'https://www.googleapis.com/auth/drive.readonly'

let cachedToken = null      // { token, expiresAt }
let inFlightToken = null    // coalesces concurrent refreshes

/** Parse and sanity-check the service account JSON from the environment. */
export function loadServiceAccount() {
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_JSON
  if (!raw) return { ok: false, reason: 'GOOGLE_SERVICE_ACCOUNT_JSON not set' }

  let sa
  try {
    sa = JSON.parse(raw)
  } catch (e) {
    return { ok: false, reason: `not valid JSON: ${e.message}` }
  }
  if (!sa.client_email) return { ok: false, reason: 'JSON has no client_email' }
  if (!sa.private_key)  return { ok: false, reason: 'JSON has no private_key' }

  // Defensive: if the value was pasted somewhere that escaped the newlines,
  // the key arrives with literal backslash-n instead of real line breaks and
  // every signature fails with an opaque error. Normalise rather than make
  // someone debug that.
  const private_key = sa.private_key.includes('\\n')
    ? sa.private_key.replace(/\\n/g, '\n')
    : sa.private_key

  if (!private_key.includes('BEGIN')) {
    return { ok: false, reason: 'private_key does not look like a PEM block' }
  }
  return { ok: true, client_email: sa.client_email, private_key, project_id: sa.project_id }
}

const b64url = (buf) =>
  Buffer.from(buf).toString('base64')
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')

/**
 * A cached OAuth access token for the service account.
 * Tokens last an hour; refresh a minute early to avoid edge-of-expiry races.
 */
export async function getAccessToken() {
  if (cachedToken && Date.now() < cachedToken.expiresAt - 60_000) {
    return cachedToken.token
  }
  // Coalesce concurrent refreshes. Without this, N requests arriving on a
  // cold or expired cache each sign their own JWT and each hit the token
  // endpoint — wasted latency on every one of them, and needless load on
  // an endpoint that rate-limits. The first caller does the work; the rest
  // await the same promise.
  if (inFlightToken) return inFlightToken
  inFlightToken = _refreshAccessToken().finally(() => { inFlightToken = null })
  return inFlightToken
}

async function _refreshAccessToken() {
  const sa = loadServiceAccount()
  if (!sa.ok) throw new Error(`Drive auth unavailable — ${sa.reason}`)

  const now = Math.floor(Date.now() / 1000)
  const header = b64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }))
  const claims = b64url(JSON.stringify({
    iss:   sa.client_email,
    scope: SCOPE,
    aud:   TOKEN_URL,
    iat:   now,
    exp:   now + 3600,
  }))
  const signer = crypto.createSign('RSA-SHA256')
  signer.update(`${header}.${claims}`)
  const signature = b64url(signer.sign(sa.private_key))
  const assertion = `${header}.${claims}.${signature}`

  const res = await fetch(TOKEN_URL, {
    method:  'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion,
    }),
  })
  const text = await res.text()
  if (!res.ok) {
    // Google's error body is the useful part here; surface it rather than
    // just a status code.
    throw new Error(`token exchange failed (${res.status}): ${text.slice(0, 300)}`)
  }
  const json = JSON.parse(text)
  cachedToken = {
    token: json.access_token,
    expiresAt: Date.now() + (json.expires_in || 3600) * 1000,
  }
  return cachedToken.token
}

/**
 * File metadata. `supportsAllDrives` is required because these documents
 * live on a Shared Drive (see the handling throughout Code.js).
 */
export async function getFileMeta(fileId) {
  const token = await getAccessToken()
  const url = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}` +
              `?fields=id,name,mimeType,size,trashed,modifiedTime&supportsAllDrives=true`
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } })
  const text = await res.text()
  if (!res.ok) throw new Error(`files.get meta failed (${res.status}): ${text.slice(0, 300)}`)
  return JSON.parse(text)
}

/**
 * Raw file bytes. Returns the undici Response so the caller can either
 * stream it straight through or buffer it — streaming is the point, since
 * the Apps Script path cannot deliver a single byte until it has base64'd
 * the entire file.
 */
export async function fetchFileResponse(fileId, { timeoutMs = 20000 } = {}) {
  const token = await getAccessToken()
  const url = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}` +
              `?alt=media&supportsAllDrives=true`

  // Bounded on purpose. Removing Apps Script's ~30-35s response ceiling
  // does NOT make Google's random 10-55s service hangs go away — it only
  // changes which ceiling we hit, and the next one along is Railway's
  // gateway timeout, which surfaces as an opaque 502 after ~30s. Failing
  // at 20s instead gives a clean error the client's auto-retry can act on,
  // and a retry is a fresh request rather than a continuation of a stall.
  //
  // This bounds time-to-first-byte, not the whole download; a large file
  // that has started streaming is making progress and is not the failure
  // mode being guarded against.
  const ctrl = new AbortController()
  const timer = setTimeout(() => ctrl.abort(), timeoutMs)
  let res
  try {
    res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
      signal: ctrl.signal,
    })
  } catch (err) {
    if (err?.name === 'AbortError') {
      throw new Error(`Drive did not respond within ${timeoutMs}ms`)
    }
    throw err
  } finally {
    clearTimeout(timer)
  }

  if (!res.ok) {
    const text = await res.text()
    throw new Error(`files.get media failed (${res.status}): ${text.slice(0, 300)}`)
  }
  return res
}

/** True when the environment is configured well enough to try Drive at all. */
export function isConfigured() {
  return loadServiceAccount().ok
}
