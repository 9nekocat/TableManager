import { useState, useEffect } from 'react'
import { useApp } from '../context/AppContext'

export default function RefsPanel() {
  const { apiFetch, setActivePanel } = useApp()
  const [tables, setTables] = useState([])
  const [target, setTarget] = useState('')
  const [sheet, setSheet] = useState('')
  const [row, setRow] = useState('')
  const [mode, setMode] = useState('query') // query | dangling
  const [data, setData] = useState(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState(null)

  useEffect(() => {
    apiFetch('/tables').then(res => {
      const files = []
      for (const cat of res) for (const t of cat.tables || []) files.push(t.file)
      setTables(files.sort())
    }).catch(() => {})
  }, [apiFetch])

  const runQuery = async () => {
    if (mode === 'query' && !target) { setError('请选择目标表'); return }
    setLoading(true); setError(null)
    try {
      let url
      if (mode === 'dangling') {
        url = '/refs/dangling' + (target ? `?file=${encodeURIComponent(target)}` : '')
      } else {
        const params = new URLSearchParams({ file: target })
        if (sheet) params.set('sheet', sheet)
        if (row) params.set('row', row)
        url = `/refs?${params}`
      }
      const res = await apiFetch(url, {}, 180000)
      setData(res)
    } catch (e) {
      setError(e.message)
      setData(null)
    } finally {
      setLoading(false)
    }
  }

  // 按来源文件分组
  const grouped = {}
  if (data?.references) {
    for (const r of data.references) {
      ;(grouped[r.src_file] = grouped[r.src_file] || []).push(r)
    }
  }

  return (
    <div className="panel">
      <div className="panel-header">
        <span className="val-title">🔗 引用查找</span>
        <span className="spacer" />
        <button className="btn" onClick={() => setActivePanel(null)} style={{ fontSize: 10 }}>关闭</button>
      </div>

      <div style={{ display: 'flex', gap: 6, padding: '8px 12px', alignItems: 'center', flexWrap: 'wrap' }}>
        <button className={`btn ${mode === 'query' ? 'btn-primary' : ''}`} style={{ fontSize: 10 }}
          onClick={() => { setMode('query'); setData(null) }}>谁引用了它</button>
        <button className={`btn ${mode === 'dangling' ? 'btn-primary' : ''}`} style={{ fontSize: 10 }}
          onClick={() => { setMode('dangling'); setData(null) }}>悬空引用</button>
        <select className="btn" style={{ fontSize: 11, maxWidth: 260 }} value={target}
          onChange={e => setTarget(e.target.value)}>
          <option value="">{mode === 'dangling' ? '（全部来源表）' : '选择目标表...'}</option>
          {tables.map(f => <option key={f} value={f}>{f}</option>)}
        </select>
        {mode === 'query' && (
          <>
            <input className="btn" style={{ fontSize: 11, width: 90 }} placeholder="sheet(可选)"
              value={sheet} onChange={e => setSheet(e.target.value)} />
            <input className="btn" style={{ fontSize: 11, width: 70 }} placeholder="行号(可选)" type="number"
              value={row} onChange={e => setRow(e.target.value)} />
          </>
        )}
        <button className="btn btn-primary" style={{ fontSize: 11 }} onClick={runQuery} disabled={loading}>
          {loading ? '查询中...' : '查询'}
        </button>
        {data && <span style={{ fontSize: 11, color: 'var(--text-dim)', fontFamily: 'var(--font-mono)' }}>{data.total} 条结果</span>}
      </div>

      {error && <div style={{ padding: '4px 12px', color: 'var(--error)', fontSize: 11 }}>{error}</div>}
      {loading && <div className="loading"><div className="spinner" /> 构建引用索引中（首次较慢）...</div>}

      {data && !loading && Object.keys(grouped).length === 0 && (
        <div className="empty-state" style={{ padding: 24 }}>
          <div style={{ fontSize: 24, opacity: 0.3, marginBottom: 8 }}>∅</div>
          <div>{mode === 'dangling' ? '没有悬空引用' : '没有表引用它'}</div>
        </div>
      )}

      {data && !loading && Object.entries(grouped).map(([src, recs]) => (
        <details key={src} open={Object.keys(grouped).length <= 3} style={{ borderBottom: '1px solid rgba(255,255,255,0.03)' }}>
          <summary className="val-file-header" style={{ outline: 'none' }}>
            <span style={{ color: 'var(--accent)', fontWeight: 600 }}>{recs.length}</span>
            <span>{src}</span>
          </summary>
          <table className="val-error-table">
            <thead><tr><th>位置</th><th>字段</th><th>类型</th><th>规则</th><th>引用值</th><th>指向</th></tr></thead>
            <tbody>
              {recs.map((r, i) => (
                <tr key={i}>
                  <td style={{ color: 'var(--warn)' }}>{r.src_sheet}!{r.src_col_ref}</td>
                  <td style={{ color: 'var(--text-dim)' }}>{r.src_header || '—'}</td>
                  <td style={{ color: r.level === 'warning' ? 'var(--warn)' : 'var(--text-dim)' }}>{r.kind}</td>
                  <td style={{ color: 'var(--text-dim)' }}>{r.rule}</td>
                  <td className="err-detail" title={r.value}>{r.value}</td>
                  <td style={{ color: 'var(--text-dim)' }}>
                    {r.targets.length ? r.targets.map(t => `${t.sheet}!A${t.row}`).join(', ') : '✗ 不存在'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </details>
      ))}
    </div>
  )
}
