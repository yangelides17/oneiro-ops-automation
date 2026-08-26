// Multi-select pill filter bar. `selected` is a Set<string>; an empty
// Set means "no filter" and the "All" pill shows active. Extracted from
// Dashboard so the WO Tracker and the Doc Status docs-queue share one
// filter control (single source of truth, guaranteed-consistent look).
const ALL = 'All'

export default function FilterBar({ label, options, selected, onToggle, onClear }) {
  const allActive = selected.size === 0
  const pillClass = (active) =>
    `text-xs px-3 py-1 rounded-full border font-medium transition-all
     ${active
       ? 'bg-navy text-white border-navy'
       : 'bg-white text-slate-600 border-slate-200 hover:border-navy/40'
     }`
  return (
    <div className="flex items-center gap-2 flex-wrap">
      <span className="text-xs font-semibold text-slate-500">{label}</span>
      <button
        key={ALL}
        onClick={onClear}
        className={pillClass(allActive)}
      >
        {ALL}
      </button>
      {options.map(opt => (
        <button
          key={opt}
          onClick={() => onToggle(opt)}
          className={pillClass(selected.has(opt))}
        >
          {opt}
        </button>
      ))}
    </div>
  )
}
