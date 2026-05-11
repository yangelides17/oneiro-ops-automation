import { useState, useEffect } from 'react'
import { Routes, Route, NavLink } from 'react-router-dom'
import Dashboard   from './pages/Dashboard'
import FieldReport from './pages/FieldReport'
import SignIn      from './pages/SignIn'
import ScanWO      from './pages/ScanWO'
import Approvals   from './pages/Approvals'
import {
  PendingCountsProvider,
  usePendingCounts,
} from './lib/PendingCountsContext'

const NAV_ITEMS = [
  { to: '/',             label: 'Dashboard',    end: true,  badgeKey: null                },
  { to: '/scan-wo',      label: 'Scan WO',      end: false, badgeKey: null                },
  { to: '/approvals',    label: 'Approvals',    end: false, badgeKey: 'approvals_review'  },
  { to: '/field-report', label: 'Field Report', end: false, badgeKey: null                },
  { to: '/sign-in',      label: 'Sign-In',      end: false, badgeKey: 'signins_pending'   },
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

function Header() {
  // Below sm: (640px) the five tabs collapse into a hamburger panel.
  // At sm+ the inline tab row stays exactly as it was on desktop.
  const [open, setOpen] = useState(false)
  const { counts } = usePendingCounts()

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

        {/* Desktop tab row (sm+) */}
        <nav className="hidden sm:flex items-center gap-1">
          {NAV_ITEMS.map(item => (
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
      </div>

      {/* Slide-down nav panel (xs only). Tapping a link auto-closes. */}
      {open && (
        <nav className="sm:hidden bg-navy border-t border-white/10
                        px-4 py-2 flex flex-col gap-1">
          {NAV_ITEMS.map(item => (
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
  return (
    <PendingCountsProvider>
      <ColdStartCounts />
      <div className="min-h-screen flex flex-col">
        <Header />
        <main className="flex-1">
          <Routes>
            <Route path="/"             element={<Dashboard />} />
            <Route path="/scan-wo"      element={<ScanWO />} />
            <Route path="/approvals"    element={<Approvals />} />
            <Route path="/field-report" element={<FieldReport />} />
            <Route path="/sign-in"      element={<SignIn />} />
          </Routes>
        </main>
        <footer className="text-center text-slate-400 text-xs py-6">
          Oneiro Collection LLC &mdash; Operations Platform
        </footer>
      </div>
    </PendingCountsProvider>
  )
}
