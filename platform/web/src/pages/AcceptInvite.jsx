import { useState } from 'react'
import { useNavigate, useSearchParams } from 'react-router-dom'
import { useAuth } from '../App'

export default function AcceptInvite() {
  const [searchParams] = useSearchParams()
  const token = searchParams.get('token') || ''
  const [form, setForm] = useState({ name: '', password: '' })
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)
  const { checkAuth } = useAuth()
  const navigate = useNavigate()

  const set = (key) => (e) => setForm(f => ({ ...f, [key]: e.target.value }))

  const handleSubmit = async (e) => {
    e.preventDefault()
    if (!token) { setError('Invalid invitation link'); return }
    setError('')
    setLoading(true)
    try {
      const r = await fetch('/api/auth/accept-invite', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ token, name: form.name, password: form.password }),
      })
      const d = await r.json()
      if (!r.ok) throw new Error(d.error || 'Failed to accept invitation')
      await checkAuth()
      navigate('/')
    } catch (err) {
      setError(err.message)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="min-h-screen flex items-center justify-center bg-navy px-4">
      <div className="max-w-sm w-full bg-white rounded-2xl shadow-xl p-8">
        <div className="text-center mb-6">
          <div className="w-12 h-12 bg-gold rounded-xl flex items-center justify-center
                          font-black text-navy text-lg mx-auto mb-4 select-none">O</div>
          <h1 className="text-navy font-black text-xl">Accept invitation</h1>
          <p className="text-slate-400 text-sm mt-1">Set up your account to join the team</p>
        </div>

        {!token ? (
          <div className="bg-amber-50 border border-amber-200 rounded-lg p-4 text-sm text-amber-700">
            This invitation link appears to be invalid or expired. Please ask your administrator
            to send a new invite.
          </div>
        ) : (
          <form onSubmit={handleSubmit} className="space-y-4">
            {error && (
              <div className="bg-red-50 border border-red-200 rounded-lg p-3 text-sm text-red-700">
                {error}
              </div>
            )}

            <div>
              <label className="field-label">Your name</label>
              <input type="text" value={form.name} onChange={set('name')}
                     className="field-input" required autoFocus />
            </div>

            <div>
              <label className="field-label">Password</label>
              <input type="password" value={form.password} onChange={set('password')}
                     className="field-input" required minLength={8} />
              <p className="text-slate-400 text-xs mt-1">At least 8 characters</p>
            </div>

            <button type="submit" disabled={loading}
                    className="btn-primary w-full">
              {loading ? 'Setting up…' : 'Create account & join'}
            </button>
          </form>
        )}
      </div>
    </div>
  )
}
