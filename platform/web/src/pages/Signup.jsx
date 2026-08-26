import { useState } from 'react'
import { Link, useNavigate } from 'react-router-dom'
import { useAuth } from '../App'

export default function Signup() {
  const [form, setForm] = useState({ name: '', email: '', password: '', orgName: '' })
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)
  const { signup } = useAuth()
  const navigate = useNavigate()

  const set = (key) => (e) => setForm(f => ({ ...f, [key]: e.target.value }))

  const handleSubmit = async (e) => {
    e.preventDefault()
    setError('')
    setLoading(true)
    try {
      await signup(form)
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
          <h1 className="text-navy font-black text-xl">Create your account</h1>
          <p className="text-slate-400 text-sm mt-1">Set up your organization</p>
        </div>

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
            <label className="field-label">Email</label>
            <input type="email" value={form.email} onChange={set('email')}
                   className="field-input" required />
          </div>

          <div>
            <label className="field-label">Password</label>
            <input type="password" value={form.password} onChange={set('password')}
                   className="field-input" required minLength={8} />
            <p className="text-slate-400 text-xs mt-1">At least 8 characters</p>
          </div>

          <div>
            <label className="field-label">Organization name</label>
            <input type="text" value={form.orgName} onChange={set('orgName')}
                   className="field-input" required placeholder="e.g. Oneiro Collection LLC" />
          </div>

          <button type="submit" disabled={loading}
                  className="btn-primary w-full">
            {loading ? 'Creating account…' : 'Create account'}
          </button>
        </form>

        <p className="text-center text-slate-400 text-sm mt-6">
          Already have an account?{' '}
          <Link to="/login" className="text-navy font-semibold hover:underline">
            Sign in
          </Link>
        </p>
      </div>
    </div>
  )
}
