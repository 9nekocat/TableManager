import { useApp } from '../context/AppContext'
import TableRow from './TableRow'

const CAT_COLORS = {
  Cards: '#65d49a', Combat: '#ff7f96', Events: '#ffc46b',
  Dialogue: '#ff9bc8', StoryMap: '#77d7ff', Metadata: '#a893ff',
  Localization: '#45b6e8', Projectiles: '#7fd8d2',
}

const SKELETON = Array.from({ length: 8 }, (_, i) => (
  <div className="skeleton-row" key={i}>
    <div className="skeleton-bar" style={{ width: '55%' }} />
    <div className="skeleton-bar" style={{ width: '55%' }} />
    <div className="skeleton-bar" style={{ width: '65%' }} />
    <div className="skeleton-bar" style={{ width: '60%' }} />
    <div className="skeleton-bar" style={{ width: '45%' }} />
    <div className="skeleton-bar" style={{ width: '40%' }} />
  </div>
))

const flattenSearchValue = (value) => {
  if (value == null) return ''
  if (Array.isArray(value)) return value.map(flattenSearchValue).join(' ')
  if (typeof value === 'object') return Object.values(value).map(flattenSearchValue).join(' ')
  return String(value)
}

const tableMatchesSearch = (table, query) => {
  const terms = query.toLocaleLowerCase().split(/\s+/).filter(Boolean)
  if (!terms.length) return true

  const haystack = [
    table.filename,
    table.file,
    table.ue_name,
    table.catKey,
    table.catDisplay,
    table.display_name,
    table.path,
    table.relative_path,
    flattenSearchValue(table.sheets),
    flattenSearchValue(table.data_sheets),
    flattenSearchValue(table.sheet_names),
  ].filter(Boolean).join(' ').toLocaleLowerCase()

  return terms.every(term => haystack.includes(term))
}

export default function TableGrid() {
  const { categories, flatTableData, activeCategory, narrowMode, searchQuery, apiFetch, showToast, refreshTables } = useApp()
  const normalizedSearch = searchQuery.trim()
  const emptyText = normalizedSearch ? `没有匹配 "${normalizedSearch}" 的表格` : '无匹配表格'

  const handleHeaderDragOver = (e) => {
    e.preventDefault()
    e.currentTarget.classList.add('drag-over')
  }
  const handleHeaderDragLeave = (e) => e.currentTarget.classList.remove('drag-over')
  const handleHeaderDrop = (e, catKey) => {
    e.preventDefault()
    e.currentTarget.classList.remove('drag-over')
    const file = e.dataTransfer.getData('text/plain')
    const fromCat = e.dataTransfer.getData('cat')
    if (file && fromCat !== catKey) {
      apiFetch('/categories/move', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ file, to_category: catKey }),
      }, 10000).then(d => {
        if (d.ok) { showToast(`已移动到新分类`, 'success'); refreshTables() }
        else showToast(`移动失败: ${d.error}`, 'error')
      }).catch(err => showToast('移动失败: ' + err.message, 'error'))
    }
  }

  if (!categories.length) {
    return <div className="content"><div className="loading">{SKELETON}</div></div>
  }

  // Narrow mode: flat list
  if (narrowMode) {
    let filtered = flatTableData
    if (activeCategory) filtered = filtered.filter(t => t.catKey === activeCategory)
    filtered = filtered.filter(t => tableMatchesSearch(t, normalizedSearch))
    if (!filtered.length) return <div className="content"><div className="empty-state">{emptyText}</div></div>

    return (
      <div className="content">
        <div className="table-grid">
          {filtered.map(t => <TableRow key={t.file} t={t} narrowMode />)}
        </div>
      </div>
    )
  }

  // Normal mode: grouped by category
  let visible = activeCategory ? categories.filter(c => c.key === activeCategory) : categories
  visible = visible
    .map(c => ({
      ...c,
      tables: c.tables.filter(t => tableMatchesSearch({ ...t, catKey: c.key, catDisplay: c.display_name }, normalizedSearch)),
    }))
    .filter(c => c.tables.length > 0)

  if (!visible.length) {
    return <div className="content"><div className="empty-state">{emptyText}</div></div>
  }

  return (
    <div className="content">
      {visible.map(c => (
        <div className="cat-section" key={c.key} data-cat={c.key}>
          <div className="cat-section-header"
            onDragOver={handleHeaderDragOver}
            onDragLeave={handleHeaderDragLeave}
            onDrop={(e) => handleHeaderDrop(e, c.key)}>
            <span className="cat-dot" style={{ background: CAT_COLORS[c.key] || '#888' }} />
            {c.display_name}
          </div>
          <div className="table-grid">
            {c.tables.length
              ? c.tables.map(t => <TableRow key={t.file} t={t} />)
              : <div className="empty-state" style={{ padding: 16 }}>此分类无表格</div>
            }
          </div>
        </div>
      ))}
    </div>
  )
}
