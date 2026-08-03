/**
 * Short-lived cache of document-batch listings, keyed by an opaque token.
 *
 * Why this exists: /api/documents/list-batch computes the listing for the
 * preview step, then /api/documents/batch-download immediately computed it
 * AGAIN from scratch before zipping. That listing is the expensive part of
 * the whole flow — an Apps Script call that walks the Drive archive — so it
 * was being paid twice per download.
 *
 * It was also a correctness bug, not just waste. The second run could return
 * a DIFFERENT set than the one the admin looked at and approved: a doc that
 * flipped Done/Sent in between, a re-archived file, or a different slice at
 * the MAX_BATCH_FILES cap. The zip, and the Sent flags, followed the second
 * set. Reusing the previewed listing makes "what you approved" and "what you
 * downloaded" the same thing by construction.
 *
 * Deliberately in-process and deliberately lossy: on any miss the caller
 * recomputes exactly as before, so an eviction, a restart, or a second
 * Railway replica costs a slow download, never a wrong one.
 */
import crypto from 'node:crypto'

const TTL_MS      = 30 * 60 * 1000   // long enough to leave the modal open
const MAX_ENTRIES = 20               // ~200KB each worst case → a few MB

const store = new Map()   // token → { filters, listing, createdAt }

/**
 * Canonical form of a filter set, for proving a download request describes
 * the same batch the preview did.
 *
 * Drops the two fields that legitimately differ between the two calls
 * (`mark_sent` is chosen at download time; `batch_token` is what we're
 * validating). Sorts object keys and the three list-valued filters, since
 * the UI builds those in click order — an admin who selects the same
 * contractors in a different order picked the same batch.
 */
export function canonicalFilters(filters) {
  const src = filters || {}
  const out = {}
  for (const key of Object.keys(src).sort()) {
    if (key === 'mark_sent' || key === 'batch_token') continue
    const v = src[key]
    out[key] = Array.isArray(v) ? [...v].map(String).sort() : v
  }
  return JSON.stringify(out)
}

/** Evict expired entries, then the oldest, until we're under the cap. */
function prune(now) {
  for (const [token, entry] of store) {
    if (now - entry.createdAt > TTL_MS) store.delete(token)
  }
  // Map preserves insertion order and entries are never re-inserted, so the
  // first key is always the oldest.
  while (store.size >= MAX_ENTRIES) {
    const oldest = store.keys().next().value
    if (oldest === undefined) break
    store.delete(oldest)
  }
}

/** Store a listing; returns the token the client should send back. */
export function putListing(filters, listing) {
  const now = Date.now()
  prune(now)
  const token = crypto.randomUUID()
  store.set(token, { filters, listing, createdAt: now })
  return token
}

/**
 * The entry for `token`, or null if it is unknown or expired.
 * Returns { filters, listing, ageMs }.
 */
export function getListing(token) {
  if (!token) return null
  const entry = store.get(token)
  if (!entry) return null
  const ageMs = Date.now() - entry.createdAt
  if (ageMs > TTL_MS) { store.delete(token); return null }
  return { filters: entry.filters, listing: entry.listing, ageMs }
}

/**
 * Resolve a download request against the cache.
 * Returns { listing, filters, ageMs } on a usable hit, else { reason }
 * explaining the miss so the caller can log why it had to recompute.
 */
export function resolveListing(token, requestFilters) {
  if (!token) return { reason: 'no-token' }
  const entry = getListing(token)
  if (!entry) return { reason: 'missing-or-expired' }
  if (canonicalFilters(entry.filters) !== canonicalFilters(requestFilters)) {
    return { reason: 'filter-mismatch' }
  }
  return { listing: entry.listing, filters: entry.filters, ageMs: entry.ageMs }
}

/** Test seam — not used in production paths. */
export function _reset() { store.clear() }
export const _internals = { TTL_MS, MAX_ENTRIES, store }
