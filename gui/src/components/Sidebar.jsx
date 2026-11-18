import { useApp } from '../context/AppContext'

const CAT_COLORS = {
  Cards: '#65d49a', Combat: '#ff7f96', Events: '#ffc46b',
  Dialogue: '#ff9bc8', StoryMap: '#77d7ff', Metadata: '#a893ff',
  Localization: '#45b6e8', Projectiles: '#7fd8d2',
}

export default function Sidebar() {
  const {
    categories,
    activeCategory,
    setActiveCategory,
    flatTableData,
    narrowMode,
    apiFetch,
    showToast,
    refreshTables,
  } = useApp()
  if (narrowMode) return null

  const totalTables = flatTableData.length
  const totalRows = flatTableData.reduce((s, t) => s + (t.row_count || 0), 0)

  return (
    <aside className="sidebar">
      <div className="sidebar-header">
        <h1>DataTable<br />Manager</h1>
        <div className="subtitle">catsintheshell</div>
      </div>

      <nav className="sidebar-nav">
        {categories.map(c => (
          <div key={c.key}
            className={`cat-item${activeCategory === c.key ? ' active' : ''}`}
            onClick={() => setActiveCategory(activeCategory === c.key ? null : c.key)}
            onDragOver={e => { e.preventDefault(); e.currentTarget.classList.add('drag-over') }}
            onDragLeave={e => e.currentTarget.classList.remove('drag-over')}
            onDrop={e => {
              e.preventDefault()
              e.currentTarget.classList.remove('drag-over')
              const file = e.dataTransfer.getData('text/plain')
              const fromCat = e.dataTransfer.getData('cat')
              if (file && fromCat !== c.key) {
                apiFetch('/categories/move', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({ file, to_category: c.key }),
                }, 10000).then(d => {
                  if (d.ok) {
                    showToast(`已移动到 ${c.display_name}`, 'success')
                    refreshTables()
                  } else {
                    showToast(`移动失败: ${d.error}`, 'error')
                  }
                }).catch(err => {
                  showToast('移动失败: ' + err.message, 'error')
                })
              }
            }}
          >
            <span className="cat-dot" style={{ background: CAT_COLORS[c.key] || '#888' }} />
            <span className="cat-label">{c.display_name}</span>
            <span className="cat-count">{c.tables.length}</span>
          </div>
        ))}
      </nav>

      <div className="sidebar-character" />

      <div className="sidebar-footer">
        <div className="info-row"><span>表格总数</span><span>{totalTables}</span></div>
        <div className="info-row"><span>数据行总计</span><span>{totalRows.toLocaleString()}</span></div>
      </div>
    </aside>
  )
}
