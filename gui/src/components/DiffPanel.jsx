import { useState } from 'react'
import { useApp } from '../context/AppContext'

export default function DiffPanel() {
  const { flatTableData, apiFetch, setActivePanel, showToast } = useApp()
  const [file, setFile] = useState('')
  const [revision, setRevision] = useState('')
  const [revList, setRevList] = useState([])
  const [result, setResult] = useState(null)
  const [loading, setLoading] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [error, setError] = useState('')

  const loadLog = async (f) => {
    setFile(f)
    setRevision('')
    setRevList([])
    setResult(null)
    if (!f) return
    try {
      const data = await apiFetch(`/diff/log/${encodeURIComponent(f)}`, {}, 15000)
      if (!data.error) setRevList(data.entries || [])
    } catch (_) {}
  }

  const runDiff = async () => {
    if (!file || !revision) return
    setLoading(true)
    setError('')
    try {
      const data = await apiFetch('/diff/compare', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ file, revision }),
      }, 60000)
      if (data.error) { setError(data.error); setResult(null) }
      else setResult(data)
    } catch (e) {
      setError(e.message)
    } finally {
      setLoading(false)
    }
  }

  const exportDiff = async () => {
    if (!file || !revision) return
    setExporting(true)
    try {
      const data = await apiFetch('/diff/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ file, revision }),
      }, 180000)
      if (data.error) showToast('导出失败: ' + data.error, 'error')
      else showToast('比对结果已在 Excel 中打开', 'success')
    } catch (e) {
      showToast('导出失败: ' + e.message, 'error')
    } finally {
      setExporting(false)
    }
  }

  const s = result?.summary
  const isColumnChange = (c) => (c.type || '').startsWith('column_')
  const badgeClass = (c) => (
    c.type === 'added' || c.type === 'column_added' ? 'add'
      : c.type === 'deleted' || c.type === 'column_deleted' ? 'del'
        : 'mod'
  )
  const badgeText = (c) => {
    if (c.type === 'added') return '+'
    if (c.type === 'deleted') return '-'
    if (c.type === 'column_added') return '+列'
    if (c.type === 'column_deleted') return '-列'
    if (c.type === 'column_modified') return '~列'
    if (c.type === 'column_moved') return '↔列'
    return '~'
  }
  const columnText = (c) => {
    if (c.type === 'column_added') return `新增列 ${c.new_name || c.rowname || ''} (new col ${c.new_col || ''})`
    if (c.type === 'column_deleted') return `删除列 ${c.old_name || c.rowname || ''} (old col ${c.old_col || ''})`
    if (c.type === 'column_modified') return `列修改 ${c.old_name || ''} → ${c.new_name || ''} (old col ${c.old_col || ''}, new col ${c.new_col || ''})`
    if (c.type === 'column_moved') return `列移动 ${c.new_name || c.old_name || c.rowname || ''} (old col ${c.old_col || ''}, new col ${c.new_col || ''})`
    return c.rowname || ''
  }

  return (
    <div className="panel">
      <div className="panel-header">
        <span>版本比对</span>
        <select value={file} onChange={e => loadLog(e.target.value)}
          style={{ background: 'var(--bg-input)', border: '1px solid var(--border)', borderRadius: 'var(--radius)', color: 'var(--text)', fontSize: 11, padding: '3px 6px', fontFamily: 'var(--font-mono)', minWidth: 180 }}>
          <option value="">-- 选择表格 --</option>
          {flatTableData.map(t => (
            <option key={t.file} value={t.file}>{t.filename}{t.is_localization ? ' (L10N)' : ''}</option>
          ))}
        </select>
        <select value={revision} onChange={e => setRevision(e.target.value)}
          disabled={!revList.length}
          style={{ background: 'var(--bg-input)', border: '1px solid var(--border)', borderRadius: 'var(--radius)', color: 'var(--text)', fontSize: 11, padding: '3px 6px', fontFamily: 'var(--font-mono)', minWidth: 200 }}>
          <option value="">-- 选择版本 --</option>
          {revList.map(e => (
            <option key={e.revision} value={e.revision}>{e.revision} | {e.author} | {(e.date || '').substring(0, 19)}</option>
          ))}
        </select>
        <button className="btn btn-primary" onClick={runDiff} disabled={!file || !revision} style={{ fontSize: 10 }}>比对</button>
        <button className="btn" onClick={exportDiff} disabled={!file || !revision || exporting} style={{ fontSize: 10 }}>
          {exporting ? '导出中...' : '导出 Excel'}
        </button>
        <span className="spacer" />
        <button className="btn" onClick={() => setActivePanel(null)} style={{ fontSize: 10 }}>关闭</button>
      </div>

      {loading && <div className="loading"><div className="spinner" /> 比对中...</div>}
      {error && <div className="empty-state" style={{ padding: 20 }}>{error}</div>}

      {result && !loading && (
        <div>
          {s && s.added === 0 && s.deleted === 0 && s.modified === 0 ? (
            <div className="empty-state" style={{ padding: 20 }}>无差异 — 文件与 {result.revision} 版本一致</div>
          ) : (
            <>
              <div style={{ padding: '6px 16px', display: 'flex', gap: 12, fontSize: 11, fontFamily: 'var(--font-mono)', borderBottom: '1px solid var(--border)' }}>
                <span>比对: <b style={{ color: 'var(--text-bright)' }}>{result.file}</b> vs <b style={{ color: 'var(--accent)' }}>{result.revision}</b></span>
                <span className="spacer" />
                {s && <><span style={{ color: 'var(--success)' }}>+{s.added} 新增</span>
                  <span style={{ color: 'var(--error)' }}>-{s.deleted} 删除</span>
                  <span style={{ color: 'var(--warn)' }}>~{s.modified} 修改</span></>}
              </div>
              {result.sheets && Object.entries(result.sheets).filter(([, sd]) => sd.status !== 'same').map(([sn, sd]) => (
                <div key={sn}>
                  <div style={{ padding: '4px 16px', fontSize: 11, fontWeight: 600, color: 'var(--text-bright)', background: 'var(--bg-row)' }}>
                    {sn} <span style={{ fontWeight: 400, color: 'var(--text-dim)', fontSize: 10 }}>{sd.changes?.length || 0} 项变更</span>
                  </div>
                  {(sd.changes || []).map((c, i) => (
                    <div key={i} className={`diff-row ${c.type === 'added' ? 'added' : c.type === 'deleted' ? 'deleted' : ''}`}>
                      <span className={`diff-badge ${badgeClass(c)}`}>
                        {badgeText(c)}
                      </span>
                      <span style={{ minWidth: 90, color: 'var(--text-bright)', fontWeight: 600, flexShrink: 0 }}>{isColumnChange(c) ? '列变更' : c.rowname}</span>
                      <span style={{ flex: 1 }}>
                        {isColumnChange(c) ? <span style={{ color: 'var(--accent)', fontSize: 10 }}>{columnText(c)}</span>
                          : c.type === 'added' ? <span style={{ color: 'var(--success)', fontSize: 10 }}>新增行</span>
                          : c.type === 'deleted' ? <span style={{ color: 'var(--error)', fontSize: 10 }}>已删除</span>
                            : (c.cell_diffs || []).map((d, j) => (
                              <span key={j} style={{ fontSize: 10, padding: '1px 5px', margin: '0 2px', borderRadius: 2, background: 'rgba(255,255,255,0.03)' }}>
                                <span style={{ color: 'var(--text-dim)' }}>col{d.col + 1}</span>
                                {' '}
                                <span style={{ color: 'var(--error)', textDecoration: 'line-through' }}>{(d.old || '').substring(0, 20)}</span>
                                <span style={{ color: 'var(--text-dim)', margin: '0 2px' }}>→</span>
                                <span style={{ color: 'var(--success)' }}>{(d.new || '').substring(0, 20)}</span>
                              </span>
                            ))}
                      </span>
                    </div>
                  ))}
                </div>
              ))}
            </>
          )}
        </div>
      )}
    </div>
  )
}
