import { createContext, useContext, useState, useCallback } from 'react'

// Shared cache of "is there work waiting?" counts feeding the nav
// badges in App.jsx and the Doc Status tab badge in Dashboard.jsx.
//
// Architecture: pages already fetch the data needed to compute their
// own count, so each page calls setCount('<key>', n) when its fetch
// resolves. App.jsx makes one cold-start fetch to /api/pending-counts
// so a fresh visitor on a non-queue page still sees nav numbers.
//
// null  = unknown — hide the badge.
// 0     = known and empty — also hide the badge.
// > 0   = render the badge.
const Ctx = createContext(null)

export function PendingCountsProvider({ children }) {
  const [counts, setCounts] = useState({
    approvals_review:      null,
    approved_docs_pending: null,
    signins_pending:       null,
    doc_status_pending:    null,
  })
  const setCount = useCallback((key, value) => {
    setCounts(prev => prev[key] === value ? prev : { ...prev, [key]: value })
  }, [])
  return (
    <Ctx.Provider value={{ counts, setCount }}>
      {children}
    </Ctx.Provider>
  )
}

export function usePendingCounts() {
  const v = useContext(Ctx)
  if (!v) throw new Error('usePendingCounts must be used inside <PendingCountsProvider>')
  return v
}
