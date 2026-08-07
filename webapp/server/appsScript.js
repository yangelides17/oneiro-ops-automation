/**
 * Shared Apps Script proxy.
 *
 * Every server-side call into the Apps Script web app goes through this
 * module. It used to be a private function inside server.js, with a
 * SECOND, simpler copy living in qb.js — which meant the redirect
 * hardening below (and all the [AS] diagnostics) covered only some of
 * our traffic, while the QuickBooks status poll that fires every 120s
 * from every open admin tab kept using bare fetch and kept hitting the
 * exact bug this module exists to fix. One implementation, one place.
 *
 * NOTE: workers/watch_and_fill.py has its own Python equivalent
 * (`_apps_script_post`) in a separate Railway dyno. It has the same
 * POST→GET-on-redirect exposure and is NOT covered by this module.
 */

import https from 'https'
import { monitorEventLoopDelay } from 'node:perf_hooks'

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

export function noteBatchDownloadStart() { batchDownloadsActive++ }
export function noteBatchDownloadEnd()   { batchDownloadsActive-- }

const ms = (ns) => Math.round(ns / 1e6)

// Single-hop HTTPS request, no auto-redirect-following. Deliberately NOT
// the global `fetch` here: fetch's `redirect: 'manual'` mode returns an
// "opaqueredirect" response per the WHATWG spec — status reads as 0 and
// headers (including Location) are unreadable, even for a same-origin
// server-to-server call. That makes it useless for actually seeing what
// Apps Script's redirect looks like. Node's raw https.request has no such
// restriction, so we get a real status + Location on every hop.
function httpRequestOnce(urlStr, { method, headers, body }) {
  return new Promise((resolve, reject) => {
    const u = new URL(urlStr)
    const req = https.request(u, { method, headers }, (res) => {
      const tHeaders = Date.now()
      const chunks = []
      res.on('data', (c) => chunks.push(c))
      res.on('end', () => resolve({
        status:   res.statusCode,
        headers:  res.headers,
        bodyText: Buffer.concat(chunks).toString('utf8'),
        tHeaders,
      }))
      res.on('error', reject)
    })
    req.on('error', reject)
    if (body) req.write(body)
    req.end()
  })
}

export async function callAppsScript(action, data = null) {
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
  // NOTE: still no timeout/retry on the ACTION itself, deliberately. A
  // blanket 28s AbortSignal was tried and reverted (644d17a) —
  // list_documents_for_batch legitimately runs longer, so the timeout
  // turned a slow-but-working call into a guaranteed failure. The one
  // retry that does exist (below) is scoped to re-fetching an
  // already-computed result, never to re-invoking doPost.
  //
  // 2026-08-06: real field incidents showed calls coming back as Google's
  // own "Script function not found: doGet" error page — the fingerprint
  // of a POST silently downgraded to a bodyless GET when Apps Script
  // answers with a 3xx (standard HTTP redirect semantics for non-GET
  // methods on 301/302/303). Added per-hop redirect logging ([AS]
  // REDIRECT) plus a targeted retry: when the chain passes through a
  // googleusercontent.com content-delivery URL (proof doPost already ran)
  // and the eventual result still isn't clean, re-fetch that same URL a
  // few times before giving up.
  //
  // Do NOT read the presence of a redirect as evidence of a problem:
  // Apps Script answers essentially EVERY web app request this way, big
  // or small, fast or slow. An earlier revision of this comment claimed
  // the redirect was specific to "slow/large" responses and that framing
  // sent one round of debugging down the wrong path. What's abnormal is
  // a chain that never RESOLVES, not a chain that exists.
  const started = Date.now()
  loopLag.reset()
  asInFlight++

  let currentUrl     = url
  let currentMethod  = 'POST'
  let currentBody    = reqJson
  let currentHeaders = {
    'Content-Type':   'application/json',
    'Content-Length': Buffer.byteLength(reqJson),
  }
  let anyRedirect = false
  let result, tHeaders
  let hop0Ms = 0   // POST /exec — the hop that carries the whole execution

  // Only these statuses are redirects fetch (and browsers) auto-follow;
  // 300/304-306/309+ are NOT.
  const REDIRECT_STATUSES = new Set([301, 302, 303, 307, 308])
  // fetch follows up to 20 redirects before giving up (WHATWG spec
  // default) and treats exceeding that as a hard network error rather
  // than silently handing back a still-3xx response. Match that.
  const MAX_HOPS = 20

  try {
    for (let hop = 0; hop < MAX_HOPS; hop++) {
      // Per-hop timing. This is the measurement that separates the two
      // candidate causes, which the aggregate `elapsed` cannot:
      //   slow hop 0 (POST /exec)  → the EXECUTION is slow/starved;
      //                              Apps Script only redirects once
      //                              doPost has finished.
      //   slow hop 1 (GET echo)    → the execution finished but Google
      //                              can't hand back the result.
      const hopStarted = Date.now()
      result = await httpRequestOnce(currentUrl, {
        method:  currentMethod,
        headers: currentHeaders,
        body:    currentMethod === 'GET' ? null : currentBody,
      })
      const hopMs = Date.now() - hopStarted
      if (hop === 0) hop0Ms = hopMs
      if (REDIRECT_STATUSES.has(result.status)) {
        anyRedirect = true
        const location = result.headers.location || ''
        // Session/cookie diagnostic. A browser following this 302 would
        // carry any cookie set on the first response; this client has
        // never sent one. If Google's content layer sometimes wants that
        // session, "go back to /exec" is exactly the bounce we see on
        // failures. Log what it hands us so we can stop guessing.
        // OBSERVE ONLY — we do not send these back. Whether we should is
        // the next question, but that would be a behavior change and the
        // point right now is to find out whether a session cookie even
        // exists to drop.
        const sc = result.headers['set-cookie']
        const cookieNames = sc
          ? (Array.isArray(sc) ? sc : [sc]).map(c => String(c).split('=')[0]).join(',')
          : 'none'
        console.warn(
          `[AS] REDIRECT action=${action} hop=${hop} status=${result.status} ` +
          `hopMs=${hopMs} fromMethod=${currentMethod} ` +
          `host=${new URL(currentUrl).host} setCookie=${cookieNames} location="${location}"`
        )
        // THE failure signature we're hunting: the content-delivery host
        // itself redirecting instead of serving. hop0 healthy + this =
        // "doPost finished and produced output, but Google won't hand it
        // back." Dump every response header here — this is the one moment
        // that might name the reason, and it is rare enough that the
        // verbosity costs nothing.
        if (new URL(currentUrl).host === 'script.googleusercontent.com') {
          console.error(
            `[AS] ECHO-BOUNCE action=${action} hop=${hop} status=${result.status} ` +
            `hopMs=${hopMs} → ${location}\n` +
            `[AS]   headers=${JSON.stringify(result.headers)}`
          )
        }
        if (!location) break
        // Reaching a googleusercontent.com/macros/echo URL means doPost
        // has already finished server-side — Apps Script only redirects
        // there once it has output to hand back. The per-hop host/location
        // logged above records that, so a failing chain can be read as
        // "execution finished, delivery failed" vs "execution never
        // produced output" without any extra bookkeeping here.
        const resolved = new URL(location, currentUrl).toString()
        currentUrl = resolved
        // Mirrors what fetch's automatic 'follow' mode does: 303 always
        // downgrades to GET; 301/302 downgrade a non-GET/HEAD method to
        // GET and drop the body; 307/308 preserve method + body.
        const downgrade = result.status === 303 ||
          ((result.status === 301 || result.status === 302) &&
            currentMethod !== 'GET' && currentMethod !== 'HEAD')
        if (downgrade) {
          currentMethod  = 'GET'
          currentBody    = null
          currentHeaders = {}
        }
        continue
      }
      console.log(
        `[AS] FINAL-HOP action=${action} hop=${hop} status=${result.status} ` +
        `hopMs=${hopMs} host=${new URL(currentUrl).host}`
      )
      break
    }

    // NO RETRY HERE, deliberately.
    //
    // An echo-URL retry lived here and was removed 2026-08-06: measured
    // zero recoveries across ~9 attempts on real incidents while adding
    // latency to failures users were already waiting through. The
    // mechanism explains why it could never have worked. Apps Script only
    // redirects to the echo URL once doPost has finished, and that URL
    // serves the execution's stored output; on the failing calls it
    // answers 302-back-to-/exec instead — i.e. there is no stored output
    // to serve. Re-fetching where absent output would live cannot succeed
    // no matter how many times we ask.
    //
    // This module is now a pure OBSERVER. It must not alter behavior
    // relative to a plain redirect-following fetch. Do not add retries,
    // timeouts, or concurrency limits here until the failure mode is
    // actually confirmed — four such fixes shipped against four unproven
    // theories and none of them stopped the recurrence.
    tHeaders = result.tHeaders
  } finally {
    asInFlight--
  }

  const text     = result.bodyText
  const finished = Date.now()
  const elapsed  = finished - started
  const ttfbMs   = tHeaders - started
  const bodyMs   = finished - tHeaders
  const elMaxMs  = ms(loopLag.max)
  const elMeanMs = Number.isFinite(loopLag.mean) ? ms(loopLag.mean) : 0
  const ctype    = result.headers['content-type'] || ''
  const clen     = result.headers['content-length'] || ''
  const trimmed  = text.trim()
  const looksHtml = trimmed.startsWith('<')
  // The chain gave up (hop cap hit, or a redirect with no Location) while
  // still sitting on an unresolved 3xx — distinct from actually landing on
  // a terminal HTML error page, and a different failure with a different
  // likely cause (platform load/instability, not auth/deployment).
  const redirectExhausted = REDIRECT_STATUSES.has(result.status)

  console.log(
    `[AS] action=${action} status=${result.status} redirected=${anyRedirect} ` +
    `elapsed=${elapsed}ms ttfb=${ttfbMs}ms body=${bodyMs}ms ` +
    `elMax=${elMaxMs}ms elMean=${elMeanMs}ms ` +
    `asInFlight=${asInFlight} zips=${batchDownloadsActive} ` +
    `reqBytes=${reqJson.length} respBytes=${text.length} ` +
    `contentLength=${clen || 'n/a'} contentType="${ctype}" ` +
    `finalUrl=${currentUrl} kind=${looksHtml ? 'HTML' : 'JSON/other'}`
  )

  if (redirectExhausted) {
    console.error(
      `[AS] REDIRECT-EXHAUSTED action=${action} status=${result.status} ` +
      `elapsed=${elapsed}ms finalUrl=${currentUrl} respBytes=${text.length}`
    )
    throw new Error(
      `Apps Script's redirect chain for "${action}" never resolved (still ${result.status} after ` +
      `following every hop, or a hop had no Location header) — this looks like Google-side ` +
      `platform load/instability, not an auth or deployment problem. Retrying in a moment is the ` +
      `right response, not re-authorizing the deployment.`
    )
  }

  // Apps Script normally returns 200 + JSON. Known causes of an HTML body
  // instead: the content-delivery echo URL failing to serve an
  // already-computed result (see the echo-retry block above), a genuine
  // Google-side error page, or — less commonly than originally assumed —
  // an actual re-auth/stale-deployment page. Don't default-blame
  // re-authorization; check the [AS] REDIRECT / HTML-RESPONSE lines above
  // for what actually happened.
  if (looksHtml) {
    // Strip tags/scripts/styles so the actual Google error message is
    // legible in the logs.
    const readable = text
      .replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<script[\s\S]*?<\/script>/gi, ' ')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&')
      .replace(/\s+/g, ' ')
      .trim()
    const titleMatch = text.match(/<title[^>]*>([\s\S]*?)<\/title>/i)
    console.error(
      `[AS] HTML-RESPONSE action=${action} status=${result.status} ` +
      `elapsed=${elapsed}ms ttfb=${ttfbMs}ms body=${bodyMs}ms ` +
      `elMax=${elMaxMs}ms zips=${batchDownloadsActive} ` +
      `respBytes=${text.length} finalUrl=${currentUrl}\n` +
      `[AS]   verdict=${elMaxMs > 1000 ? 'node-blocked (event loop stalled)' : 'apps-script-slow'}\n` +
      `[AS]   title="${titleMatch ? titleMatch[1].trim() : '(none)'}"\n` +
      `[AS]   readable="${readable.slice(0, 600)}"`
    )
    const status = result.status
    const redirected = anyRedirect
      ? ' (redirect chain didn\'t deliver a clean response — check [AS] REDIRECT lines above for the actual cause)'
      : ''
    throw new Error(
      `Apps Script returned HTML instead of JSON${redirected} — ` +
      `action="${action}", status=${status}.`
    )
  }
  let json
  try {
    json = JSON.parse(text)
  } catch (e) {
    console.error(
      `[AS] NON-JSON action=${action} status=${result.status} respBytes=${text.length} ` +
      `head="${text.slice(0, 200).replace(/\n/g, ' ')}"`
    )
    throw new Error(
      `Apps Script response wasn't JSON (action="${action}", status=${result.status}): ` +
      `${text.slice(0, 200)}`
    )
  }
  // Server-side timing rider from jsonResponse_. Read BEFORE the
  // json.error throw below so failures are measured too.
  //
  // `loaded`/`entry`/`exit` all come from Apps Script's clock, so only
  // same-clock deltas are used — no client/server skew enters the math.
  // The remainder is what nothing could see before: how much of the call
  // is spent before our code even starts running.
  if (json && json._t && typeof json._t.exit === 'number') {
    const { loaded, entry, exit } = json._t
    const dispatchMs = entry - loaded      // rest of module eval + dispatch
    const handlerMs  = exit  - entry       // actual handler work
    const inScriptMs = exit  - loaded      // everything we can attribute
    const preScriptMs = hop0Ms - inScriptMs // network + queue + parse/compile
    console.log(
      `[AS] SPLIT action=${action} hop0=${hop0Ms}ms ` +
      `preScript=${preScriptMs}ms dispatch=${dispatchMs}ms handler=${handlerMs}ms ` +
      `(preScript = network + queue + fetch/parse/compile of Code.js)`
    )
    // Slow/failed Drive ops inside that execution, when there were any.
    // Testing whether _withDriveRetry_'s 3-attempt budget is what turns a
    // ~10s Drive hiccup into a >30s execution that blows Apps Script's
    // response ceiling and fails the whole request.
    if (Array.isArray(json._t.drive) && json._t.drive.length) {
      console.warn(
        `[AS] DRIVE-OPS action=${action} handler=${handlerMs}ms ` +
        json._t.drive.map(d => `${d.op}#${d.n}=${d.ms}ms/${d.ok ? 'ok' : 'FAIL'}`).join(' ')
      )
    }
  }
  if (json.error) {
    console.error(`[AS] JSON-ERROR action=${action} status=${result.status} error="${json.error}"`)
    throw new Error(json.error)
  }
  return json
}
