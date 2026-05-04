import { useState } from 'react'
import { Routes, Route, NavLink } from 'react-router-dom'
import Dashboard   from './pages/Dashboard'
import FieldReport from './pages/FieldReport'
import SignIn      from './pages/SignIn'
import ScanWO      from './pages/ScanWO'
import Approvals   from './pages/Approvals'

const NAV_ITEMS = [
  { to: '/',             label: 'Dashboard',    end: true  },
  { to: '/scan-wo',      label: 'Scan WO',      end: false },
  { to: '/approvals',    label: 'Approvals',    end: false },
  { to: '/field-report', label: 'Field Report', end: false },
  { to: '/sign-in',      label: 'Sign-In',      end: false },
]

function Header() {
  // Below sm: (640px) the five tabs collapse into a hamburger panel.
  // At sm+ the inline tab row stays exactly as it was on desktop.
  const [open, setOpen] = useState(false)

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
              {item.label}
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
              {item.label}
            </NavLink>
          ))}
        </nav>
      )}
    </header>
  )
}

export default function App() {
  return (
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
  )
}
