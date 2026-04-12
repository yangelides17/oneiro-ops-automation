import { Routes, Route, NavLink, useLocation } from 'react-router-dom'
import Dashboard   from './pages/Dashboard'
import FieldReport from './pages/FieldReport'

function Header() {
  const loc = useLocation()
  const onField = loc.pathname === '/field-report'

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
          <span className="text-white/50 text-[11px] block leading-none mt-0.5 hidden sm:block">
            Oneiro Collection LLC
          </span>
        </div>

        {/* Nav */}
        <nav className="flex items-center gap-1">
          <NavLink
            to="/"
            end
            className={({ isActive }) =>
              `nav-link ${isActive ? 'active' : ''}`
            }
          >
            Dashboard
          </NavLink>
          <NavLink
            to="/field-report"
            className={({ isActive }) =>
              `nav-link ${isActive ? 'active' : ''}`
            }
          >
            Field Report
          </NavLink>
        </nav>
      </div>
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
          <Route path="/field-report" element={<FieldReport />} />
        </Routes>
      </main>
      <footer className="text-center text-slate-400 text-xs py-6">
        Oneiro Collection LLC &mdash; Operations Platform
      </footer>
    </div>
  )
}
