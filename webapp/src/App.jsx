import { useState, useEffect } from 'react'
import { Routes, Route, NavLink, Navigate } from 'react-router-dom'
import Dashboard   from './pages/Dashboard'
import NavTab      from './pages/NavTab'
import FieldReport from './pages/FieldReport'
import SignIn      from './pages/SignIn'
import ScanWO      from './pages/ScanWO'
import Approvals   from './pages/Approvals'
import {
  PendingCountsProvider,
  usePendingCounts,
} from './lib/PendingCountsContext'

// Access gate. The site is locked behind a shared "door code" carried in
// the user's link (`/?key=<code>`): one code grants 'admin' (full app),
// another grants 'crew' (Nav + Field Report + Sign-In only). The code is
// validated SERVER-SIDE (see server.js "Access gate") — it never lives in
// this bundle — and the resolved role is what the client renders against.
//
// Role values: 'admin' | 'crew' | null (locked) | 'loading' (undecided).
// Rotating a code in Railway instantly locks out old links on next load.
const ACCESS_LOADING = 'loading'

function useAccess() {
  const [role, setRole] = useState(ACCESS_LOADING)

  useEffect(() => {
    if (typeof window === 'undefined') return
    let cancelled = false

    const params = new URLSearchParams(window.location.search)
    const key = params.get('key')

    // Strip ?key= from the address bar so the code isn't left sitting in
    // the URL / browser history after it's been redeemed.
    const stripKey = () => {
      params.delete('key')
      const qs = params.toString()
      const clean = window.location.pathname + (qs ? '?' + qs : '') + window.location.hash
      window.history.replaceState(null, '', clean)
    }

    async function resolve() {
      try {
        if (key) {
          // Redeem the code: server validates + sets the access cookie.
          const r = await fetch('/api/access/login', {
            method:  'POST',
            headers: { 'Content-Type': 'application/json' },
            body:    JSON.stringify({ key }),
          })
          stripKey()
          const d = await r.json().catch(() => ({}))
          if (!cancelled) setRole(d.role || null)
          return
        }
        // No code in the link — check for an existing valid cookie.
        const r = await fetch('/api/access/session')
        const d = await r.json().catch(() => ({}))
        if (!cancelled) setRole(d.role || null)
      } catch {
        // Network error: fail closed to the locked screen rather than
        // flashing the app to someone who may not be authorized.
        if (!cancelled) setRole(null)
      }
    }
    resolve()
    return () => { cancelled = true }
  }, [])

  return role
}

// Full-screen gate shown while the role is undecided ('loading') or when
// access is denied (null). Deliberately gives nothing away beyond "use
// your link" — no hint that a code exists or what it looks like.
function AccessGate({ locked }) {
  return (
    <div className="min-h-screen flex items-center justify-center bg-navy px-4">
      <div className="max-w-sm w-full bg-white rounded-2xl shadow-xl p-8 text-center">
        <div className="w-12 h-12 bg-gold rounded-xl flex items-center justify-center
                        font-black text-navy text-lg mx-auto mb-4 select-none">O</div>
        {locked ? (
          <>
            <h1 className="text-navy font-black text-lg">Access restricted</h1>
            <p className="text-slate-500 text-sm mt-2 leading-relaxed">
              This site is private. Please open it using the access link
              your team shared with you.
            </p>
            <p className="text-slate-400 text-xs mt-4">
              If you believe you should have access, contact your Oneiro Ops administrator.
            </p>
          </>
        ) : (
          <>
            <p className="text-navy font-bold">Oneiro Ops</p>
            <p className="text-slate-400 text-sm mt-2">Checking access…</p>
          </>
        )}
      </div>
    </div>
  )
}

// Hook: catch the ?qb=connected | ?qb=error&msg=... query params left
// behind by the OAuth callback redirect. Surfaces a brief toast and
// strips the params from the URL so a refresh doesn't re-show them.
function useQbAuthResult() {
  const [toast, setToast] = useState(null)  // { kind: 'success'|'error', text }

  useEffect(() => {
    if (typeof window === 'undefined') return
    const params = new URLSearchParams(window.location.search)
    const qb = params.get('qb')
    if (!qb) return

    if (qb === 'connected') {
      setToast({ kind: 'success', text: 'QuickBooks connected.' })
    } else if (qb === 'error') {
      setToast({ kind: 'error', text: 'QuickBooks connection failed: ' + (params.get('msg') || 'unknown error') })
    }

    // Strip qb + msg from URL
    params.delete('qb')
    params.delete('msg')
    const qs = params.toString()
    const clean = window.location.pathname + (qs ? '?' + qs : '') + window.location.hash
    window.history.replaceState(null, '', clean)

    // Auto-dismiss success after 5s; errors stick until manually dismissed
    const t = setTimeout(() => setToast(prev => prev?.kind === 'success' ? null : prev), 5000)
    return () => clearTimeout(t)
  }, [])

  return { toast, dismissToast: () => setToast(null) }
}

function QbAuthToast({ toast, onDismiss }) {
  if (!toast) return null
  const isError = toast.kind === 'error'
  return (
    <div className={`fixed top-4 right-4 z-50 max-w-sm rounded-xl shadow-lg p-4 flex items-start gap-3 ${
      isError ? 'bg-red-50 border border-red-200' : 'bg-emerald-50 border border-emerald-200'
    }`}>
      <span className={`text-xl flex-shrink-0 leading-none ${isError ? 'text-red-500' : 'text-emerald-500'}`}>
        {isError ? '✕' : '✓'}
      </span>
      <div className="flex-1 text-sm">
        <p className={`font-bold ${isError ? 'text-red-800' : 'text-emerald-800'}`}>
          {toast.text}
        </p>
      </div>
      <button onClick={onDismiss}
              className={`text-xs font-bold leading-none ${isError ? 'text-red-600 hover:text-red-800' : 'text-emerald-600 hover:text-emerald-800'}`}
              aria-label="Dismiss">×</button>
    </div>
  )
}

// `roles` gates which links each role sees. Crew is limited to Nav,
// Field Report, and Sign-In; admins see everything. Keep this in sync
// with the <Route> table in App() below.
const NAV_ITEMS = [
  { to: '/',             label: 'Dashboard',    end: true,  badgeKey: null,               roles: ['admin']         },
  { to: '/nav',          label: 'Nav',          end: false, badgeKey: null,               roles: ['admin', 'crew'] },
  { to: '/scan-wo',      label: 'Scan WO',      end: false, badgeKey: null,               roles: ['admin']         },
  { to: '/approvals',    label: 'Approvals',    end: false, badgeKey: 'approvals_review', roles: ['admin']         },
  { to: '/field-report', label: 'Field Report', end: false, badgeKey: null,               roles: ['admin', 'crew'] },
  { to: '/sign-in',      label: 'Sign-In',      end: false, badgeKey: 'signins_pending',  roles: ['admin', 'crew'] },
]

// Subtle amber pill matching the visual weight of the existing nav.
// Hidden when the count is null (unknown) or 0 (known-empty) so the
// nav reads exactly like today when there's nothing to act on.
export function NavBadge({ n }) {
  if (n == null || n === 0) return null
  return (
    <span className="ml-1.5 inline-flex items-center justify-center
                     min-w-[18px] h-[18px] px-1 rounded-full
                     text-[10px] font-bold bg-amber-500 text-white
                     align-middle leading-none">
      {n > 99 ? '99+' : n}
    </span>
  )
}

function Header({ role = 'admin' }) {
  // Below sm: (640px) the tabs collapse into a hamburger panel. At sm+
  // the inline tab row stays as it was on desktop. The link set is
  // filtered by role, so crews see only their allowed pages (Nav, Field
  // Report, Sign-In) and never a link into Dashboard / Approvals / etc.
  const [open, setOpen] = useState(false)
  const { counts } = usePendingCounts()
  const items = NAV_ITEMS.filter(item => item.roles.includes(role))

  // Render label + optional badge. badgeKey is null for nav items that
  // shouldn't surface a count (Dashboard / Scan WO / Field Report).
  const linkContent = (item) => (
    <>
      {item.label}
      {item.badgeKey && <NavBadge n={counts[item.badgeKey]} />}
    </>
  )

  return (
    <header className="bg-navy sticky top-0 z-50 shadow-lg">
      <div className="max-w-6xl mx-auto px-4 h-14 flex items-center gap-3">
        {/* Logo mark */}
        <div className="w-8 h-8 bg-gold rounded-lg flex items-center justify-center
                        font-black text-navy text-sm flex-shrink-0 select-none">
          O
        </div>

        {/* Brand */}
        <div className="flex-1 min-w-0">
          <span className="text-white font-bold text-[15px] leading-none">Oneiro Ops</span>
          <span className="text-white/50 text-[11px] leading-none mt-0.5 hidden sm:block">
            Oneiro Collection LLC
          </span>
        </div>

        {(
          <>
            {/* Desktop tab row (sm+) */}
            <nav className="hidden sm:flex items-center gap-1">
              {items.map(item => (
                <NavLink
                  key={item.to}
                  to={item.to}
                  end={item.end}
                  className={({ isActive }) => `nav-link ${isActive ? 'active' : ''}`}
                >
                  {linkContent(item)}
                </NavLink>
              ))}
            </nav>

            {/* Hamburger toggle (xs only) */}
            <button
              type="button"
              aria-label={open ? 'Close navigation' : 'Open navigation'}
              aria-expanded={open}
              onClick={() => setOpen(o => !o)}
              className="sm:hidden text-white p-2 -mr-2 rounded-lg
                         hover:bg-white/10 active:bg-white/15 transition-colors"
            >
              {open ? (
                <svg width="22" height="22" viewBox="0 0 24 24" fill="none"
                     stroke="currentColor" strokeWidth="2.2" strokeLinecap="round">
                  <line x1="6"  y1="6"  x2="18" y2="18" />
                  <line x1="18" y1="6"  x2="6"  y2="18" />
                </svg>
              ) : (
                <svg width="22" height="22" viewBox="0 0 24 24" fill="none"
                     stroke="currentColor" strokeWidth="2.2" strokeLinecap="round">
                  <line x1="4" y1="7"  x2="20" y2="7"  />
                  <line x1="4" y1="12" x2="20" y2="12" />
                  <line x1="4" y1="17" x2="20" y2="17" />
                </svg>
              )}
            </button>
          </>
        )}
      </div>

      {/* Slide-down nav panel (xs only). Tapping a link auto-closes. */}
      {open && (
        <nav className="sm:hidden bg-navy border-t border-white/10
                        px-4 py-2 flex flex-col gap-1">
          {items.map(item => (
            <NavLink
              key={item.to}
              to={item.to}
              end={item.end}
              onClick={() => setOpen(false)}
              className={({ isActive }) => `nav-link ${isActive ? 'active' : ''}`}
            >
              {linkContent(item)}
            </NavLink>
          ))}
        </nav>
      )}
    </header>
  )
}

// Cold-start fetches: populate all four nav-badge counts so a fresh
// visitor on any page sees pending work in the nav. Two requests fire
// in parallel:
//   1. /api/pending-counts (fast, ~300 ms) — Approvals + Sign-In +
//      Approved Docs counts. Drives the always-visible nav badges.
//   2. /api/pending-counts/doc-status (slower, ~500 ms-1.5 s) — Doc
//      Status pending count. Drives the Doc Status tab badge. Slow
//      because it runs the same _buildDocStatusPayload_ the calendar
//      uses; split out so the fast badges don't wait on it.
// Pages with their own queue fetches will still overwrite their
// matching slot afterwards — no polling.
function ColdStartCounts() {
  const { setCount } = usePendingCounts()
  useEffect(() => {
    let cancelled = false

    fetch('/api/pending-counts')
      .then(r => r.ok ? r.json() : Promise.reject(r))
      .then(d => {
        if (cancelled) return
        if (d?.approvals_review      !== undefined) setCount('approvals_review',      d.approvals_review)
        if (d?.approved_docs_pending !== undefined) setCount('approved_docs_pending', d.approved_docs_pending)
        if (d?.signins_pending       !== undefined) setCount('signins_pending',       d.signins_pending)
      })
      .catch(err => {
        console.warn('cold-start /api/pending-counts failed:', err)
      })

    fetch('/api/pending-counts/doc-status')
      .then(r => r.ok ? r.json() : Promise.reject(r))
      .then(d => {
        if (cancelled) return
        if (d?.doc_status_pending !== undefined) {
          setCount('doc_status_pending', d.doc_status_pending)
        }
      })
      .catch(err => {
        // Non-fatal — the DocStatusTab will populate this slot when
        // the user visits Dashboard, same as before.
        console.warn('cold-start /api/pending-counts/doc-status failed:', err)
      })

    return () => { cancelled = true }
  }, [setCount])
  return null
}

export default function App() {
  const role = useAccess()
  const { toast, dismissToast } = useQbAuthResult()

  // Gate: undecided → spinner card; denied → locked card. Only a
  // resolved 'admin' | 'crew' renders the app. Hooks above run first so
  // the early return doesn't violate the rules of hooks.
  if (role === ACCESS_LOADING) return <AccessGate locked={false} />
  if (!role)                   return <AccessGate locked={true} />

  const crewMode = role === 'crew'
  return (
    <PendingCountsProvider>
      <ColdStartCounts />
      <QbAuthToast toast={toast} onDismiss={dismissToast} />
      <div className="min-h-screen flex flex-col">
        <Header role={role} />
        <main className="flex-1">
          <Routes>
            {crewMode ? (
              // Crew: only Nav, Field Report, and Sign-In are reachable.
              // Everything else bounces to Field Report. `replace` keeps
              // the back button from yo-yo'ing into a blocked route.
              <>
                <Route path="/nav"          element={<NavTab />} />
                <Route path="/field-report" element={<FieldReport />} />
                <Route path="/sign-in"      element={<SignIn />} />
                <Route path="*"             element={<Navigate to="/field-report" replace />} />
              </>
            ) : (
              <>
                <Route path="/"             element={<Dashboard />} />
                <Route path="/nav"          element={<NavTab />} />
                <Route path="/scan-wo"      element={<ScanWO />} />
                <Route path="/approvals"    element={<Approvals />} />
                <Route path="/field-report" element={<FieldReport />} />
                <Route path="/sign-in"      element={<SignIn />} />
              </>
            )}
          </Routes>
        </main>
        <footer className="text-center text-slate-400 text-xs py-6 space-y-1">
          <div>Oneiro Collection LLC &mdash; Operations Platform</div>
          <div>
            <a href="/legal/privacy.html" target="_blank" rel="noopener noreferrer"
               className="hover:text-slate-600 hover:underline">Privacy</a>
            <span className="mx-2 text-slate-300">·</span>
            <a href="/legal/eula.html" target="_blank" rel="noopener noreferrer"
               className="hover:text-slate-600 hover:underline">Terms</a>
          </div>
        </footer>
      </div>
    </PendingCountsProvider>
  )
}
