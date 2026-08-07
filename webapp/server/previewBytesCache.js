/**
 * In-process cache of approval-document PDF bytes.
 *
 * Why this exists: fetching one preview is the single worst call in the app.
 * Measured 2026-08-07, it had a 21% failure rate — the highest of anything we
 * probed — because it drags a multi-MB Drive blob through Apps Script, which
 * base64-inflates it ~33% into a JSON string. Failure rate scales with how much
 * Google service work an execution does, and this does the most.
 *
 * Yet the same bytes were re-fetched constantly and never reused:
 *   - the Approvals page auto-selects the first queued doc on mount, so every
 *     page load pulled a full PDF
 *   - clicking back to a doc you already looked at re-downloaded all of it
 *   - nothing cached it anywhere: the browser is sent `no-store` (an Intuit
 *     security requirement), Node had no cache, and Apps Script's CacheService
 *     caps entries at 100KB so a ~5MB payload could never live there
 *
 * Caching server-side fixes all three without touching the browser's no-store
 * posture — the bytes still aren't persisted on the client, we just stop asking
 * Google for something we already have.
 *
 * Bounded by TOTAL BYTES rather than entry count, because entries here are ~4MB
 * each rather than the ~200KB of batchListingCache. Deliberately in-process and
 * lossy: any miss just costs what the old code cost every single time.
 *
 * Correctness rule: a document's bytes can change underneath us (Regenerate,
 * Reupload). Those are OUR mutations and we know exactly when they happen, so
 * callers MUST call `invalidate(fileId)` on those paths. That is cheaper and
 * more reliable than polling Drive for an md5 — and staleness here would be
 * user-visible (showing the pre-regenerate PDF), so it is the main risk.
 */

const TTL_MS = Number(process.env.PREVIEW_CACHE_TTL_MS) || 30 * 60 * 1000

// Conservative by default — Railway containers are small and these entries are
// multi-MB. ~128MB holds roughly 30 typical documents, which comfortably covers
// a review session, while leaving the rest of the heap for pdf-lib and zip
// streaming (both of which are already memory-hungry on this server).
const MAX_BYTES = Number(process.env.PREVIEW_CACHE_MAX_BYTES) || 128 * 1024 * 1024

const store = new Map()   // key → { buf, mime, filename, createdAt, bytes }
let totalBytes = 0

const keyFor = (fileId, version) => `${fileId}::${version || ''}`

/** Drop an entry and keep the byte counter honest. */
function drop(key) {
  const entry = store.get(key)
  if (!entry) return
  totalBytes -= entry.bytes
  store.delete(key)
}

/** Evict expired entries, then oldest-first until we're under the byte cap. */
function prune(now) {
  for (const [key, entry] of store) {
    if (now - entry.createdAt > TTL_MS) drop(key)
  }
  // Map preserves insertion order, and `get` re-inserts on hit, so the first
  // key is always the least-recently-used.
  while (totalBytes > MAX_BYTES && store.size) {
    const oldest = store.keys().next().value
    if (oldest === undefined) break
    drop(oldest)
  }
}

/**
 * Cached bytes for a file, or null. On a hit the entry is re-inserted so
 * insertion order doubles as LRU order.
 */
export function getBytes(fileId, version) {
  const key = keyFor(fileId, version)
  const entry = store.get(key)
  if (!entry) return null
  if (Date.now() - entry.createdAt > TTL_MS) { drop(key); return null }
  store.delete(key)
  store.set(key, entry)
  return { buf: entry.buf, mime: entry.mime, filename: entry.filename, ageMs: Date.now() - entry.createdAt }
}

/** Cache a file's bytes. Oversized entries are skipped, never cached partially. */
export function putBytes(fileId, version, buf, mime, filename) {
  if (!Buffer.isBuffer(buf) || !buf.length) return
  // A single file bigger than the whole budget would evict everything and then
  // still not fit — skip it rather than thrash.
  if (buf.length > MAX_BYTES) return
  const key = keyFor(fileId, version)
  drop(key)
  store.set(key, {
    buf, mime, filename,
    createdAt: Date.now(),
    bytes: buf.length,
  })
  totalBytes += buf.length
  prune(Date.now())
}

/**
 * Forget every cached version of a file. Call this wherever the document's
 * bytes are changed server-side — regenerate, reupload, approve — so a preview
 * can never show a stale copy.
 */
export function invalidate(fileId) {
  if (!fileId) return
  const prefix = `${fileId}::`
  for (const key of [...store.keys()]) {
    if (key.startsWith(prefix)) drop(key)
  }
}

/** Snapshot for logging. */
export function stats() {
  return {
    entries: store.size,
    mb: +(totalBytes / 1024 / 1024).toFixed(1),
    capMb: +(MAX_BYTES / 1024 / 1024).toFixed(0),
  }
}

/** Test seam — not used in production paths. */
export function _reset() { store.clear(); totalBytes = 0 }
export const _internals = { TTL_MS, MAX_BYTES, store }
