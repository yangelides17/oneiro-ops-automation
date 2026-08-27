import { useState, useEffect, useCallback } from 'react'
import PdfViewer from '../components/PdfViewer'
import PdfThumbnail from '../components/PdfThumbnail'

// ─── Type Badges ──────────────────────────────────────────────

const TYPE_BADGES = {
  scan:              { label: 'Scan',   bg: 'bg-blue-100',   text: 'text-blue-700',   icon: '📄' },
  field_report:      { label: 'CFR',    bg: 'bg-green-100',  text: 'text-green-700',  icon: '📄' },
  production_log:    { label: 'PL',     bg: 'bg-amber-100',  text: 'text-amber-700',  icon: '📄' },
  signin:            { label: 'SI',     bg: 'bg-purple-100', text: 'text-purple-700', icon: '📄' },
  certified_payroll: { label: 'CP',     bg: 'bg-orange-100', text: 'text-orange-700', icon: '📄' },
  photo:             { label: 'Photo',  bg: 'bg-pink-100',   text: 'text-pink-700',   icon: '🖼' },
}

function formatDate(d) {
  if (!d) return ''
  try { return new Date(d).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) }
  catch { return d }
}

// ─── Component ────────────────────────────────────────────────

export default function FileBrowser() {
  // Navigation path: array of breadcrumb segments
  // [] = root, [contractor] = L1, [contractor, contract] = L2, [..., wo] = L3
  const [path, setPath] = useState([])
  const [items, setItems] = useState([])
  const [woFiles, setWoFiles] = useState(null) // Level 3: { scan, documents, photos }
  const [loading, setLoading] = useState(false)
  const [search, setSearch] = useState('')
  const [searchResults, setSearchResults] = useState(null)

  // Split view
  const [selectedFile, setSelectedFile] = useState(null)

  const level = path.length // 0=root, 1=contracts, 2=WOs, 3=files

  // ── Fetch data when path changes ────────────────────────────
  useEffect(() => {
    if (searchResults) return // search mode overrides navigation
    setLoading(true)
    setItems([])
    setWoFiles(null)
    setSelectedFile(null)

    const params = new URLSearchParams()

    if (level >= 1) params.set('contractorId', path[0].id)
    if (level >= 2) {
      params.set('contractNum', path[1].contractNum)
      params.set('regionCode', path[1].regionCode)
    }

    if (level === 3) {
      // Level 3: fetch files for the WO
      const woId = path[2].id
      fetch(`/api/wos/${encodeURIComponent(woId)}/files`)
        .then(r => r.json())
        .then(d => { setWoFiles(d); setLoading(false) })
        .catch(() => setLoading(false))
      return
    }

    // Levels 0-2: fetch hierarchy
    fetch(`/api/files/browse?${params}`)
      .then(r => r.json())
      .then(d => { setItems(d.items || []); setLoading(false) })
      .catch(() => setLoading(false))
  }, [path, level, searchResults])

  // ── Search ──────────────────────────────────────────────────
  useEffect(() => {
    if (!search.trim()) { setSearchResults(null); return }
    const timer = setTimeout(() => {
      fetch(`/api/files?search=${encodeURIComponent(search.trim())}&limit=20`)
        .then(r => r.json())
        .then(d => setSearchResults(d.files || []))
        .catch(() => setSearchResults([]))
    }, 300)
    return () => clearTimeout(timer)
  }, [search])

  // ── Navigation helpers ──────────────────────────────────────
  const navigateTo = useCallback((segment) => {
    setSearch('')
    setSearchResults(null)
    setSelectedFile(null)
    setPath(prev => [...prev, segment])
  }, [])

  const navigateToLevel = useCallback((idx) => {
    setSearch('')
    setSearchResults(null)
    setSelectedFile(null)
    if (idx < 0) setPath([])
    else setPath(prev => prev.slice(0, idx + 1))
  }, [])

  const clearSearch = useCallback(() => {
    setSearch('')
    setSearchResults(null)
  }, [])

  // ── Build file tiles for Level 3 ────────────────────────────
  const fileTiles = woFiles ? [
    ...(woFiles.scan ? [{
      id: 'scan',
      type: 'scan',
      label: woFiles.scan.filename || 'WO Scan',
      sublabel: 'Original scan',
      url: woFiles.scan.url,
      mimeType: 'application/pdf',
    }] : []),
    ...(woFiles.documents || []).map(d => ({
      id: d.id,
      type: 'document',
      docType: d.docType,
      label: d.filename || `${d.docType}.pdf`,
      sublabel: formatDate(d.anchorDate),
      url: d.url,
      mimeType: 'application/pdf',
    })),
    ...(woFiles.photos || []).map(p => ({
      id: p.id,
      type: 'photo',
      label: p.filename || 'Photo',
      sublabel: formatDate(p.createdAt),
      url: p.url,
      mimeType: p.mimeType || 'image/jpeg',
    })),
  ] : []

  // ── Render ──────────────────────────────────────────────────
  const showSplit = !!selectedFile
  const isImage = selectedFile?.mimeType?.startsWith('image/')

  return (
    <div className="max-w-7xl mx-auto px-4 py-5">
      {/* Top bar: breadcrumb + search */}
      <div className="flex items-center justify-between gap-3 mb-4">
        <nav className="flex items-center gap-1 text-sm min-w-0 overflow-hidden">
          <button
            type="button"
            onClick={() => navigateToLevel(-1)}
            className={`font-bold flex-shrink-0 ${level === 0 && !searchResults ? 'text-navy' : 'text-slate-400 hover:text-navy'}`}
          >
            📁 Files
          </button>
          {path.map((seg, i) => (
            <span key={i} className="flex items-center gap-1 min-w-0">
              <span className="text-slate-300 flex-shrink-0">›</span>
              <button
                type="button"
                onClick={() => navigateToLevel(i)}
                className={`truncate ${i === path.length - 1 && !searchResults ? 'font-bold text-navy' : 'text-slate-400 hover:text-navy'}`}
              >
                {seg.label}
              </button>
            </span>
          ))}
          {searchResults && (
            <span className="flex items-center gap-1">
              <span className="text-slate-300">›</span>
              <span className="font-bold text-navy">Search: "{search}"</span>
            </span>
          )}
        </nav>

        <div className="relative flex-shrink-0">
          <input
            type="text"
            value={search}
            onChange={e => setSearch(e.target.value)}
            placeholder="Search WO #…"
            className="field-input text-sm w-44 pl-8"
          />
          <span className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400 text-xs">🔍</span>
          {search && (
            <button
              type="button"
              onClick={clearSearch}
              className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 text-xs"
            >
              ✕
            </button>
          )}
        </div>
      </div>

      {/* Main content area */}
      <div className={showSplit ? 'flex gap-4' : ''}>
        {/* Left panel: tiles */}
        <div className={showSplit ? 'w-[35%] min-w-[280px] flex-shrink-0 overflow-y-auto' : 'w-full'}
             style={showSplit ? { maxHeight: 'calc(100vh - 120px)' } : undefined}>

          {loading && (
            <div className="flex items-center justify-center py-12">
              <div className="w-6 h-6 border-2 border-slate-200 border-t-navy rounded-full animate-spin" />
            </div>
          )}

          {/* Search results */}
          {searchResults && !loading && (
            <div className="space-y-1">
              {searchResults.length === 0 && (
                <p className="text-sm text-slate-400 italic py-8 text-center">No files found for "{search}"</p>
              )}
              {searchResults.map(file => (
                <SearchResultRow
                  key={`${file.type}-${file.id}`}
                  file={file}
                  selected={selectedFile?.id === file.id}
                  onClick={() => setSelectedFile(selectedFile?.id === file.id ? null : file)}
                />
              ))}
            </div>
          )}

          {/* Folder tiles (Levels 0-2) */}
          {!searchResults && !loading && level < 3 && (
            <div className={`grid gap-3 ${showSplit ? 'grid-cols-1' : 'grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 xl:grid-cols-5'}`}>
              {items.length === 0 && (
                <p className="col-span-full text-sm text-slate-400 italic py-8 text-center">No files yet</p>
              )}
              {items.map(item => (
                <FolderTile
                  key={item.id || `${item.contractNum}-${item.regionCode}`}
                  item={item}
                  level={level}
                  onClick={() => {
                    if (level === 0) navigateTo({ id: item.id, label: item.label })
                    else if (level === 1) navigateTo({ contractNum: item.contractNum, regionCode: item.regionCode, label: `${item.label} - ${item.sublabel}` })
                    else if (level === 2) navigateTo({ id: item.id, woNumber: item.woNumber, label: item.label })
                  }}
                />
              ))}
            </div>
          )}

          {/* File tiles (Level 3) */}
          {!searchResults && !loading && level === 3 && (
            <div className={`grid gap-3 ${showSplit ? 'grid-cols-1' : 'grid-cols-2 sm:grid-cols-3 lg:grid-cols-4'}`}>
              {fileTiles.length === 0 && (
                <p className="col-span-full text-sm text-slate-400 italic py-8 text-center">No files for this work order</p>
              )}
              {fileTiles.map(file => (
                <FileTile
                  key={file.id}
                  file={file}
                  selected={selectedFile?.id === file.id}
                  compact={showSplit}
                  onClick={() => setSelectedFile(selectedFile?.id === file.id ? null : file)}
                />
              ))}
            </div>
          )}
        </div>

        {/* Right panel: preview */}
        {showSplit && (
          <div className="flex-1 min-w-0" style={{ maxHeight: 'calc(100vh - 120px)' }}>
            {isImage ? (
              <div className="border border-slate-200 rounded-xl bg-white overflow-hidden h-full">
                <div className="flex items-center justify-between px-3 py-2 bg-slate-50 border-b border-slate-200">
                  <span className="text-sm font-semibold text-slate-700 truncate">
                    {selectedFile.label || selectedFile.filename}
                  </span>
                  <button
                    type="button"
                    onClick={() => setSelectedFile(null)}
                    className="w-6 h-6 flex items-center justify-center rounded-full text-slate-400 hover:text-red-500 hover:bg-red-50"
                  >×</button>
                </div>
                <div className="flex items-center justify-center bg-slate-100 p-4 overflow-auto" style={{ height: 'calc(100% - 44px)' }}>
                  <img src={selectedFile.url} alt={selectedFile.label} className="max-w-full max-h-full object-contain rounded" />
                </div>
              </div>
            ) : (
              <PdfViewer
                url={selectedFile.url}
                filename={selectedFile.label || selectedFile.filename}
                collapsed={false}
                onToggle={() => {}}
                onClose={() => setSelectedFile(null)}
              />
            )}
          </div>
        )}
      </div>
    </div>
  )
}

// ─── Folder Tile ──────────────────────────────────────────────

function FolderTile({ item, level, onClick }) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="text-left rounded-xl border border-slate-200 bg-white p-4
                 hover:bg-slate-50 hover:border-navy/30 transition-all cursor-pointer
                 flex items-start gap-3"
    >
      <span className="text-2xl flex-shrink-0 mt-0.5">📂</span>
      <div className="min-w-0 flex-1">
        <p className="text-sm font-bold text-slate-800 truncate">{item.label}</p>
        {item.sublabel && (
          <p className="text-xs text-slate-500 truncate mt-0.5">{item.sublabel}</p>
        )}
        <p className="text-[11px] text-slate-400 mt-1">
          {level < 2
            ? `${item.count} work order${item.count === 1 ? '' : 's'}`
            : `${item.count} file${item.count === 1 ? '' : 's'}`
          }
        </p>
      </div>
    </button>
  )
}

// ─── File Tile ────────────────────────────────────────────────

function FileTile({ file, selected, compact, onClick }) {
  const badge = TYPE_BADGES[file.docType || file.type] || TYPE_BADGES.scan
  const isPdf = file.mimeType === 'application/pdf'
  const isImage = file.mimeType?.startsWith('image/')

  return (
    <button
      type="button"
      onClick={onClick}
      className={`text-left rounded-xl border transition-all cursor-pointer overflow-hidden
                  ${selected
                    ? 'border-navy bg-navy/5 ring-1 ring-navy/20'
                    : 'border-slate-200 bg-white hover:bg-slate-50 hover:border-navy/30'}`}
    >
      {/* Thumbnail preview */}
      {!compact && isPdf && file.url && (
        <div className="border-b border-slate-100">
          <PdfThumbnail url={file.url} height={140} className="w-full" />
        </div>
      )}
      {!compact && isImage && file.url && (
        <div className="border-b border-slate-100 bg-slate-100">
          <img src={file.url} alt={file.label} className="w-full h-36 object-cover" />
        </div>
      )}

      {/* Label bar */}
      <div className="flex items-center gap-2 p-3">
        <span className="text-base flex-shrink-0">{badge.icon}</span>
        <div className="min-w-0 flex-1">
          <p className="text-sm font-semibold text-slate-700 truncate">{file.label}</p>
          {file.sublabel && (
            <p className="text-xs text-slate-400 mt-0.5">{file.sublabel}</p>
          )}
        </div>
        <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full flex-shrink-0 ${badge.bg} ${badge.text}`}>
          {badge.label}
        </span>
      </div>
    </button>
  )
}

// ─── Search Result Row ────────────────────────────────────────

function SearchResultRow({ file, selected, onClick }) {
  const badge = TYPE_BADGES[file.docType || file.type] || TYPE_BADGES.scan

  return (
    <button
      type="button"
      onClick={onClick}
      className={`w-full text-left flex items-center gap-3 px-3 py-2.5 rounded-xl border transition-all
                  ${selected
                    ? 'border-navy bg-navy/5 ring-1 ring-navy/20'
                    : 'border-slate-200 bg-white hover:bg-slate-50'}`}
    >
      <span className="text-base flex-shrink-0">{badge.icon}</span>
      <div className="min-w-0 flex-1">
        <p className="text-sm font-semibold text-slate-700 truncate">{file.filename}</p>
        <p className="text-xs text-slate-400">{file.woNumber} · {file.contractorName} · {formatDate(file.date)}</p>
      </div>
      <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full flex-shrink-0 ${badge.bg} ${badge.text}`}>
        {badge.label}
      </span>
    </button>
  )
}
