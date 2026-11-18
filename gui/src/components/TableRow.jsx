import { useState } from 'react'
import { useApp } from '../context/AppContext'

export default function TableRow({ t, narrowMode }) {
  const { apiFetch, showToast, setActivePanel } = useApp()
  const [diffing, setDiffing] = useState(false)
  const [dropdownOpen, setDropdownOpen] = useState(false)

  const handleDragStart = (e) => {
    e.dataTransfer.setData('text/plain', t.file)
    e.dataTransfer.setData('cat', t.catKey || '')
    e.dataTransfer.effectAllowed = 'move'
    e.currentTarget.classList.add('dragging')
  }

  const handleDragEnd = (e) => e.currentTarget.classList.remove('dragging')

  const openTable = (e) => {
    e.stopPropagation()
    apiFetch(`/open/${encodeURIComponent(t.file)}`, { method: 'POST' }, 10000)
  }

  const quickDiff = async (e) => {
    e.stopPropagation()
    setDropdownOpen(false)
    setDiffing(true)
    try {
      const data = await apiFetch(`/diff/smart-export/${encodeURIComponent(t.file)}`, {}, 60000)
      if (data.error) { showToast(data.error, 'error') }
      else showToast('比对结果已在 Excel 中打开', 'success')
    } catch (err) {
      showToast('比对失败: ' + err.message, 'error')
    } finally {
      setDiffing(false)
    }
  }

  const toggleDropdown = (e) => {
    e.stopPropagation()
    setDropdownOpen(prev => !prev)
  }

  const quickDiffInApp = (e) => {
    e.stopPropagation()
    setDropdownOpen(false)
    setActivePanel('diff')
    showToast('已打开版本比对面板，请选择版本后比对', 'success')
  }

  if (narrowMode) {
    return (
      <div className="table-row" data-file={t.file} data-cat={t.catKey || ''}
        draggable onDragStart={handleDragStart} onDragEnd={handleDragEnd}
        onClick={openTable}>
        <span className="col-name">
          <span className="file-icon">&#x1F4C4;</span>
          {t.filename}
          {t.is_localization && <span className="tag tag-local">L10N</span>}
        </span>
        <span className="col-dim">{t.row_count || 0}r &times; {t.col_count || 0}c</span>
        <span className="col-actions">
          <button className="btn" onClick={openTable}>&#x1F4C2; 打开</button>
          <button className="btn" onClick={quickDiff} disabled={diffing}>
            {diffing ? <><span className="spinner" style={{width:10,height:10,borderWidth:2}} /> 比对中...</> : <>&#x1F504; 比对</>}
          </button>
        </span>
      </div>
    )
  }

  return (
    <div className="table-row" data-file={t.file} data-cat={t.catKey || ''}
      draggable onDragStart={handleDragStart} onDragEnd={handleDragEnd}
      onClick={openTable}>
      <span className="col-name">
        <span className="file-icon">&#x1F4C4;</span>
        {t.filename}
        {t.is_localization && <span className="tag tag-local">L10N</span>}
      </span>
      <span className="col-dim">{t.row_count || 0}r &times; {t.col_count || 0}c</span>
      <span className="col-size">{t.size_fmt}</span>
      <span className="col-ue" title={t.ue_name}>{t.ue_name || '—'}</span>
      <span className="col-date">{t.last_modified}</span>
      <span className="col-actions">
        <button className="btn" onClick={openTable}>&#x1F4C2; 打开</button>
        <span className="split-btn-group">
          <button className="btn" onClick={quickDiff} disabled={diffing}>
            {diffing ? <><span className="spinner" style={{width:10,height:10,borderWidth:2}} /> 比对中...</> : <>&#x1F504; 比对</>}
          </button>
          <button className="btn split-toggle" onClick={toggleDropdown}>&#x25BE;</button>
          <div className={`split-dropdown${dropdownOpen ? ' open' : ''}`} onClick={e => e.stopPropagation()}>
            <div className="split-dropdown-item" onClick={quickDiffInApp}>应用内查看</div>
          </div>
        </span>
      </span>
    </div>
  )
}
