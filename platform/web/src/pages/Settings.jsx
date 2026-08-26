import { useState, useEffect } from 'react'
import { useAuth } from '../App'

const TABS = [
  { key: 'org',           label: 'Organization' },
  { key: 'users',         label: 'Users' },
  { key: 'regions',       label: 'Regions' },
  { key: 'contractors',   label: 'Contractors' },
  { key: 'employees',     label: 'Employees' },
  { key: 'categories',    label: 'Categories' },
  { key: 'pricing',       label: 'Pricing' },
  { key: 'payroll',       label: 'Payroll' },
  { key: 'integrations',  label: 'Integrations' },
]

function useFetch(url) {
  const [data, setData] = useState(null)
  const [loading, setLoading] = useState(true)
  const refresh = () => {
    setLoading(true)
    fetch(url).then(r => r.json()).then(setData).finally(() => setLoading(false))
  }
  useEffect(refresh, [url])
  return { data, loading, refresh }
}

// ── Organization Tab ─────────────────────────────────────────
function OrgTab() {
  const [org, setOrg] = useState(null)
  const [saving, setSaving] = useState(false)
  const [msg, setMsg] = useState('')

  useEffect(() => {
    fetch('/api/settings/org').then(r => r.json()).then(setOrg)
  }, [])

  const save = async () => {
    setSaving(true)
    setMsg('')
    const r = await fetch('/api/settings/org', {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(org),
    })
    if (r.ok) setMsg('Saved')
    else setMsg('Error saving')
    setSaving(false)
  }

  if (!org) return <p className="text-slate-400 p-4">Loading…</p>

  const field = (label, key, type = 'text') => (
    <div>
      <label className="field-label">{label}</label>
      <input type={type} value={org[key] || ''} onChange={e => setOrg({ ...org, [key]: e.target.value })}
             className="field-input" />
    </div>
  )

  return (
    <div className="space-y-4">
      {field('Organization Name', 'name')}
      {field('Address', 'address')}
      {field('Phone', 'phone')}
      {field('Email', 'email', 'email')}
      {field('Tax ID (EIN)', 'taxId')}
      {field('Timezone', 'timezone')}
      <div>
        <label className="field-label">Operational Day Cutoff Hour</label>
        <input type="number" min="0" max="23" value={org.opDayCutoffHour ?? 5}
               onChange={e => setOrg({ ...org, opDayCutoffHour: parseInt(e.target.value) })}
               className="field-input w-24" />
        <p className="text-slate-400 text-xs mt-1">Hours before this time bucket to the previous day (for night shifts)</p>
      </div>
      {field('Signatory Name', 'signatoryName')}
      {field('Signatory Title', 'signatoryTitle')}
      <div className="flex items-center gap-3">
        <button onClick={save} disabled={saving} className="btn-primary">
          {saving ? 'Saving…' : 'Save'}
        </button>
        {msg && <span className="text-sm text-emerald-600">{msg}</span>}
      </div>
    </div>
  )
}

// ── Generic CRUD List Tab ────────────────────────────────────
function CrudListTab({ endpoint, columns, newItemDefaults }) {
  const { data, loading, refresh } = useFetch(endpoint)
  const [adding, setAdding] = useState(false)
  const [newItem, setNewItem] = useState(newItemDefaults)
  const [saving, setSaving] = useState(false)

  const items = Array.isArray(data) ? data : []

  const handleAdd = async () => {
    setSaving(true)
    const r = await fetch(endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(newItem),
    })
    if (r.ok) {
      setAdding(false)
      setNewItem(newItemDefaults)
      refresh()
    }
    setSaving(false)
  }

  const handleDelete = async (id) => {
    if (!confirm('Delete this item?')) return
    await fetch(`${endpoint}/${id}`, { method: 'DELETE' })
    refresh()
  }

  return (
    <div>
      <div className="flex items-center justify-between mb-4">
        <span className="text-sm text-slate-500">{items.length} items</span>
        <button onClick={() => setAdding(true)} className="btn-ghost text-sm">
          + Add
        </button>
      </div>

      {adding && (
        <div className="card p-4 mb-4 space-y-3">
          {columns.map(col => (
            <div key={col.key}>
              <label className="field-label">{col.label}</label>
              <input type={col.type || 'text'}
                     value={newItem[col.key] || ''}
                     onChange={e => setNewItem({ ...newItem, [col.key]: e.target.value })}
                     className="field-input" />
            </div>
          ))}
          <div className="flex gap-2">
            <button onClick={handleAdd} disabled={saving} className="btn-primary text-sm">
              {saving ? 'Saving…' : 'Save'}
            </button>
            <button onClick={() => setAdding(false)} className="btn-outline text-sm">Cancel</button>
          </div>
        </div>
      )}

      {loading ? (
        <p className="text-slate-400">Loading…</p>
      ) : items.length === 0 ? (
        <p className="text-slate-400 text-sm">No items yet. Click "+ Add" to create one.</p>
      ) : (
        <div className="space-y-2">
          {items.map(item => (
            <div key={item.id} className="card px-4 py-3 flex items-center justify-between">
              <div>
                {columns.map(col => (
                  <span key={col.key} className="mr-4 text-sm">
                    <span className="text-slate-400 text-xs uppercase mr-1">{col.label}:</span>
                    {String(item[col.key] || '')}
                  </span>
                ))}
              </div>
              <button onClick={() => handleDelete(item.id)}
                      className="text-red-400 hover:text-red-600 text-xs">
                Delete
              </button>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

// ── Users Tab ────────────────────────────────────────────────
function UsersTab() {
  const { data, loading, refresh } = useFetch('/api/settings/users')
  const [inviteEmail, setInviteEmail] = useState('')
  const [inviteRole, setInviteRole] = useState('crew')
  const [inviting, setInviting] = useState(false)
  const [msg, setMsg] = useState('')

  const users = Array.isArray(data) ? data : []

  const handleInvite = async () => {
    setInviting(true)
    setMsg('')
    const r = await fetch('/api/settings/users/invite', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email: inviteEmail, role: inviteRole }),
    })
    if (r.ok) {
      setMsg(`Invitation sent to ${inviteEmail}`)
      setInviteEmail('')
      refresh()
    } else {
      const d = await r.json().catch(() => ({}))
      setMsg(d.error || 'Failed to send invite')
    }
    setInviting(false)
  }

  return (
    <div>
      <div className="card p-4 mb-6">
        <h3 className="section-label">Invite team member</h3>
        <div className="flex gap-3 items-end">
          <div className="flex-1">
            <label className="field-label">Email</label>
            <input type="email" value={inviteEmail} onChange={e => setInviteEmail(e.target.value)}
                   className="field-input" placeholder="colleague@company.com" />
          </div>
          <div>
            <label className="field-label">Role</label>
            <select value={inviteRole} onChange={e => setInviteRole(e.target.value)}
                    className="field-input">
              <option value="admin">Admin</option>
              <option value="foreman">Foreman</option>
              <option value="crew">Crew</option>
            </select>
          </div>
          <button onClick={handleInvite} disabled={inviting || !inviteEmail}
                  className="btn-primary text-sm whitespace-nowrap">
            {inviting ? 'Sending…' : 'Send invite'}
          </button>
        </div>
        {msg && <p className="text-sm text-emerald-600 mt-2">{msg}</p>}
      </div>

      <h3 className="section-label">Team members</h3>
      {loading ? <p className="text-slate-400">Loading…</p> : (
        <div className="space-y-2">
          {users.map(u => (
            <div key={u.id} className="card px-4 py-3 flex items-center justify-between">
              <div>
                <span className="font-semibold text-sm text-navy">{u.name}</span>
                <span className="text-slate-400 text-sm ml-2">{u.email}</span>
              </div>
              <span className="text-xs font-bold uppercase tracking-wider text-slate-500 bg-slate-100 px-2 py-1 rounded">
                {u.role}
              </span>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

// ── Main Settings Page ───────────────────────────────────────
export default function Settings() {
  const [activeTab, setActiveTab] = useState('org')
  const { user } = useAuth()
  const isOwner = user?.role === 'owner'

  return (
    <div className="max-w-4xl mx-auto px-4 py-6">
      <h1 className="text-navy font-black text-xl mb-6">Settings</h1>

      <div className="flex gap-2 overflow-x-auto pb-2 mb-6">
        {TABS.map(tab => (
          <button key={tab.key}
                  onClick={() => setActiveTab(tab.key)}
                  className={`px-3 py-1.5 rounded-lg text-sm font-medium whitespace-nowrap transition-colors ${
                    activeTab === tab.key
                      ? 'bg-navy text-white'
                      : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                  }`}>
            {tab.label}
          </button>
        ))}
      </div>

      <div className="card p-6">
        {activeTab === 'org' && <OrgTab />}
        {activeTab === 'users' && <UsersTab />}
        {activeTab === 'regions' && (
          <CrudListTab endpoint="/api/settings/regions"
                       columns={[{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }]}
                       newItemDefaults={{ code: '', name: '' }} />
        )}
        {activeTab === 'contractors' && (
          <CrudListTab endpoint="/api/settings/contractors"
                       columns={[
                         { key: 'name', label: 'Name' },
                         { key: 'contactEmail', label: 'Email', type: 'email' },
                       ]}
                       newItemDefaults={{ name: '', contactEmail: '' }} />
        )}
        {activeTab === 'employees' && (
          <CrudListTab endpoint="/api/settings/employees"
                       columns={[{ key: 'name', label: 'Name' }, { key: 'address', label: 'Address' }]}
                       newItemDefaults={{ name: '', address: '' }} />
        )}
        {activeTab === 'categories' && (
          <CrudListTab endpoint="/api/settings/categories"
                       columns={[
                         { key: 'name', label: 'Name' },
                         { key: 'unit', label: 'Unit' },
                         { key: 'pricingGroup', label: 'Pricing Group' },
                       ]}
                       newItemDefaults={{ name: '', unit: 'LF', pricingGroup: 'line4' }} />
        )}
        {activeTab === 'pricing' && (
          <CrudListTab endpoint="/api/settings/pricing"
                       columns={[
                         { key: 'contractNum', label: 'Contract #' },
                         { key: 'regionCode', label: 'Region' },
                         { key: 'rateLine4', label: '4" Rate' },
                       ]}
                       newItemDefaults={{ contractorId: '', contractNum: '', regionCode: '' }} />
        )}
        {activeTab === 'payroll' && (
          <div className="space-y-6">
            <div>
              <h3 className="section-label">Classifications</h3>
              <CrudListTab endpoint="/api/settings/payroll/classifications"
                           columns={[{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }]}
                           newItemDefaults={{ code: '', name: '' }} />
            </div>
            <div>
              <h3 className="section-label">Pay Rates</h3>
              <CrudListTab endpoint="/api/settings/payroll/rates"
                           columns={[
                             { key: 'classificationCode', label: 'Class' },
                             { key: 'effectiveDate', label: 'Effective', type: 'date' },
                             { key: 'rateSt', label: 'ST Rate' },
                             { key: 'rateOt', label: 'OT Rate' },
                           ]}
                           newItemDefaults={{ classificationCode: '', effectiveDate: '', rateSt: '', rateOt: '' }} />
            </div>
          </div>
        )}
        {activeTab === 'integrations' && (
          <div className="space-y-4">
            <p className="text-slate-500 text-sm">
              Integrations allow you to connect external services to your account.
            </p>
            <div className="card p-4 flex items-center justify-between">
              <div>
                <h3 className="font-semibold text-navy">Google Drive</h3>
                <p className="text-slate-400 text-xs">Sync documents to your Google Drive</p>
              </div>
              <span className="text-xs font-bold uppercase text-slate-400 bg-slate-100 px-2 py-1 rounded">
                Coming soon
              </span>
            </div>
            <div className="card p-4 flex items-center justify-between">
              <div>
                <h3 className="font-semibold text-navy">QuickBooks Online</h3>
                <p className="text-slate-400 text-xs">Generate invoices in QuickBooks</p>
              </div>
              <span className="text-xs font-bold uppercase text-slate-400 bg-slate-100 px-2 py-1 rounded">
                Coming soon
              </span>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}
