const STYLES = {
  'received':    'bg-blue-50   text-blue-700   border-blue-200',
  'dispatched':  'bg-amber-50  text-amber-700  border-amber-200',
  'in progress': 'bg-orange-50 text-orange-700 border-orange-200',
  'complete':    'bg-green-50  text-green-700  border-green-200',
}

export default function StatusBadge({ status }) {
  const key   = (status || '').toLowerCase()
  const style = STYLES[key] ?? 'bg-slate-100 text-slate-600 border-slate-200'
  const label = status || 'Unknown'

  return (
    <span className={`inline-block px-2 py-0.5 rounded-full text-[10px] font-bold
                      uppercase tracking-wider border ${style}`}>
      {label}
    </span>
  )
}
