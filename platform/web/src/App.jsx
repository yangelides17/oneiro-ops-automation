import { useState, useEffect, createContext, useContext } from 'react'
import { Routes, Route, NavLink, Navigate, useNavigate } from 'react-router-dom'
import Dashboard   from './pages/Dashboard'
import NavTab      from './pages/NavTab'
import FieldReport from './pages/FieldReport'
import SignIn      from './pages/SignIn'
import ScanWO      from './pages/ScanWO'
import FileBrowser from './pages/FileBrowser'
import Approvals   from './pages/Approvals'
import Login       from './pages/Login'
import Signup      from './pages/Signup'
import AcceptInvite from './pages/AcceptInvite'
import Settings    from './pages/Settings'
import WOAssignment from './pages/WOAssignment'
import MyWork      from './pages/MyWork'
import {
  PendingCountsProvider,
  usePendingCounts,
} from './lib/PendingCountsContext'

const AuthContext = createContext(null)
export const useAuth = () => useContext(AuthContext)

function AuthProvider({ children }) {
  const [state, setState] = useState({ loading: true, user: null, org: null })

  const checkAuth = async () => {
    try {
      const r = await fetch('/api/auth/me')
      if (!r.ok) { setState({ loading: false, user: null, org: null }); return }
      const d = await r.json()
      setState({ loading: false, user: d.user, org: d.org })
    } catch {
      setState({ loading: false, user: null, org: null })
    }
  }

  useEffect(() => { checkAuth() }, [])

  const login = async (email, password) => {
    const r = await fetch('/api/auth/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, password }),
    })
    const d = await r.json()
    if (!r.ok) throw new Error(d.error || 'Login failed')
    await checkAuth()
    return d
  }

  const signup = async (data) => {
    const r = await fetch('/api/auth/signup', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data),
    })
    const d = await r.json()
    if (!r.ok) throw new Error(d.error || 'Signup failed')
    await checkAuth()
    return d
  }

  const logout = async () => {
    await fetch('/api/auth/logout', { method: 'POST' })
    setState({ loading: false, user: null, org: null })
  }

  return (
    <AuthContext.Provider value={{ ...state, login, signup, logout, checkAuth }}>
      {children}
    </AuthContext.Provider>
  )
}

function isAdmin(role) { return role === 'owner' || role === 'admin' }

const NAV_ITEMS = [
  { to: '/',             label: 'Dashboard',    end: true,  badgeKey: null,               roles: ['owner', 'admin']                      },
  { to: '/nav',          label: 'Nav',          end: false, badgeKey: null,               roles: ['owner', 'admin', 'foreman', 'crew']   },
  { to: '/files',        label: 'Files',        end: false, badgeKey: null,               roles: ['owner', 'admin', 'foreman', 'crew']   },
  { to: '/my-work',      label: 'My Work',      end: false, badgeKey: null,               roles: ['owner', 'admin', 'foreman', 'crew']   },
  { to: '/assign',       label: 'Assign WOs',  end: false, badgeKey: null,               roles: ['owner', 'admin']                      },
  { to: '/scan-wo',      label: 'Scan WO',      end: false, badgeKey: null,               roles: ['owner', 'admin', 'foreman']           },
  { to: '/approvals',    label: 'Approvals',    end: false, badgeKey: 'approvals_review', roles: ['owner', 'admin']                      },
  { to: '/field-report', label: 'Field Report', end: false, badgeKey: null,               roles: ['owner', 'admin', 'foreman', 'crew']   },
  { to: '/sign-in',      label: 'Sign-In',      end: false, badgeKey: 'signins_pending',  roles: ['owner', 'admin', 'foreman', 'crew']   },
  { to: '/settings',     label: 'Settings',     end: false, badgeKey: null,               roles: ['owner', 'admin']                      },
]

export function NavBadge({ n }) {
  if (n == null || n === 0) return null
  return (
    <span className="ml-1.5 inline-flex items-center justify-center min-w-[18px] h-[18px] px-1 rounded-full text-[10px] font-bold bg-amber-500 text-white align-middle leading-none">
      {n > 99 ? '99+' : n}
    </span>
  )
}

function Header() {
  const { user, org, logout } = useAuth()
  const [open, setOpen] = useState(false)
  const { counts } = usePendingCounts()
  const items = NAV_ITEMS.filter(item => item.roles.includes(user?.role))
  const navigate = useNavigate()
  const linkContent = (item) => (<>{item.label}{item.badgeKey && <NavBadge n={counts[item.badgeKey]} />}</>)
  const handleLogout = async () => { await logout(); navigate('/login') }

  return (
    <header className="bg-navy sticky top-0 z-50 shadow-lg">
      <div className="max-w-6xl mx-auto px-4 h-14 flex items-center gap-3">
        <div className="w-8 h-8 bg-gold rounded-lg flex items-center justify-center font-black text-navy text-sm flex-shrink-0 select-none">O</div>
        <div className="flex-1 min-w-0">
          <span className="text-white font-bold text-[15px] leading-none">{org?.name || 'Oneiro Ops'}</span>
          <span className="text-white/50 text-[11px] leading-none mt-0.5 hidden sm:block">{user?.name} · {user?.role}</span>
        </div>
        <nav className="hidden sm:flex items-center gap-1">
          {items.map(item => (
            <NavLink key={item.to} to={item.to} end={item.end} className={({ isActive }) => `nav-link ${isActive ? 'active' : ''}`}>
              {linkContent(item)}
            </NavLink>
          ))}
        </nav>
        <button onClick={handleLogout} className="hidden sm:block text-white/50 text-xs hover:text-white transition-colors">Log out</button>
        <button type="button" aria-label={open ? 'Close navigation' : 'Open navigation'} aria-expanded={open}
          onClick={() => setOpen(o => !o)}
          className="sm:hidden text-white p-2 -mr-2 rounded-lg hover:bg-white/10 active:bg-white/15 transition-colors">
          {open
            ? <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round"><line x1="6" y1="6" x2="18" y2="18" /><line x1="18" y1="6" x2="6" y2="18" /></svg>
            : <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round"><line x1="4" y1="7" x2="20" y2="7" /><line x1="4" y1="12" x2="20" y2="12" /><line x1="4" y1="17" x2="20" y2="17" /></svg>}
        </button>
      </div>
      {open && (
        <nav className="sm:hidden bg-navy border-t border-white/10 px-4 py-2 flex flex-col gap-1">
          {items.map(item => (
            <NavLink key={item.to} to={item.to} end={item.end} onClick={() => setOpen(false)} className={({ isActive }) => `nav-link ${isActive ? 'active' : ''}`}>
              {linkContent(item)}
            </NavLink>
          ))}
          <button onClick={handleLogout} className="nav-link text-left text-white/50 hover:text-white mt-2">Log out</button>
        </nav>
      )}
    </header>
  )
}

function ColdStartCounts() {
  const { setCount } = usePendingCounts()
  useEffect(() => {
    let cancelled = false
    fetch('/api/pending-counts').then(r => r.ok ? r.json() : Promise.reject(r)).then(d => {
      if (cancelled) return
      if (d?.approvals_review !== undefined) setCount('approvals_review', d.approvals_review)
      if (d?.approved_docs_pending !== undefined) setCount('approved_docs_pending', d.approved_docs_pending)
      if (d?.signins_pending !== undefined) setCount('signins_pending', d.signins_pending)
    }).catch(() => {})
    fetch('/api/pending-counts/doc-status').then(r => r.ok ? r.json() : Promise.reject(r)).then(d => {
      if (cancelled) return
      if (d?.doc_status_pending !== undefined) setCount('doc_status_pending', d.doc_status_pending)
    }).catch(() => {})
    return () => { cancelled = true }
  }, [setCount])
  return null
}

function AuthenticatedApp() {
  const { user } = useAuth()
  const role = user?.role || 'crew'
  return (
    <PendingCountsProvider>
      <ColdStartCounts />
      <div className="min-h-screen flex flex-col">
        <Header />
        <main className="flex-1">
          <Routes>
            {isAdmin(role) ? (
              <>
                <Route path="/" element={<Dashboard />} />
                <Route path="/nav" element={<NavTab />} />
                <Route path="/files" element={<FileBrowser />} />
                <Route path="/my-work" element={<MyWork />} />
                <Route path="/assign" element={<WOAssignment />} />
                <Route path="/scan-wo" element={<ScanWO />} />
                <Route path="/approvals" element={<Approvals />} />
                <Route path="/field-report" element={<FieldReport />} />
                <Route path="/sign-in" element={<SignIn />} />
                <Route path="/settings" element={<Settings />} />
              </>
            ) : role === 'foreman' ? (
              <>
                <Route path="/nav" element={<NavTab />} />
                <Route path="/files" element={<FileBrowser />} />
                <Route path="/my-work" element={<MyWork />} />
                <Route path="/scan-wo" element={<ScanWO />} />
                <Route path="/field-report" element={<FieldReport />} />
                <Route path="/sign-in" element={<SignIn />} />
                <Route path="*" element={<Navigate to="/my-work" replace />} />
              </>
            ) : (
              <>
                <Route path="/nav" element={<NavTab />} />
                <Route path="/files" element={<FileBrowser />} />
                <Route path="/my-work" element={<MyWork />} />
                <Route path="/field-report" element={<FieldReport />} />
                <Route path="/sign-in" element={<SignIn />} />
                <Route path="*" element={<Navigate to="/my-work" replace />} />
              </>
            )}
          </Routes>
        </main>
        <footer className="text-center text-slate-400 text-xs py-6 space-y-1">
          <div>{user?.org?.name || 'Oneiro Platform'} &mdash; Operations Platform</div>
          <div>
            <a href="/legal/privacy.html" target="_blank" rel="noopener noreferrer" className="hover:text-slate-600 hover:underline">Privacy</a>
            <span className="mx-2 text-slate-300">·</span>
            <a href="/legal/eula.html" target="_blank" rel="noopener noreferrer" className="hover:text-slate-600 hover:underline">Terms</a>
          </div>
        </footer>
      </div>
    </PendingCountsProvider>
  )
}

export default function App() {
  const { loading, user } = useAuth()
  if (loading) return (
    <div className="min-h-screen flex items-center justify-center bg-navy">
      <div className="max-w-sm w-full bg-white rounded-2xl shadow-xl p-8 text-center">
        <div className="w-12 h-12 bg-gold rounded-xl flex items-center justify-center font-black text-navy text-lg mx-auto mb-4 select-none">O</div>
        <p className="text-navy font-bold">Oneiro Ops</p>
        <p className="text-slate-400 text-sm mt-2">Loading…</p>
      </div>
    </div>
  )
  if (!user) return (
    <Routes>
      <Route path="/login" element={<Login />} />
      <Route path="/signup" element={<Signup />} />
      <Route path="/accept-invite" element={<AcceptInvite />} />
      <Route path="*" element={<Navigate to="/login" replace />} />
    </Routes>
  )
  return <AuthenticatedApp />
}

export { AuthProvider }
