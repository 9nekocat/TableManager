import { useState } from 'react'
import { useApp } from '../context/AppContext'

export default function ValidationPanel() {
  const { apiFetch, setActivePanel } = useApp()
  const [mode, setMode] = useState('raw') // raw=配置表检查 | processed=导出数据检查
  const [data, setData] = useState(null)
  const [loading, setLoading] = useState(false)

  const runCheck = async (m = mode) => {
    setLoading(true)
    try {
      const res = m === 'processed'
        ? await apiFetch('/pcheck', {}, 120000)
        : await apiFetch('/check', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({}) }, 120000)
      setData(res)
    } catch (e) {
      setData({ total_errors: -1, files_affected: 0, error: e.message })
    } finally {
      setLoading(false)
    }
  }

  const switchMode = (m) => {
    if (m === mode) return
    setMode(m)
    setData(null)
    runCheck(m)
  }

  // Run on first open
  if (!data && !loading) runCheck()

  return (
    <div className="panel">
      <div className="panel-header">
        <span className="val-title" style={{ color: data?.total_errors ? 'var(--error)' : 'var(--success)' }}>
          {data?.total_errors === 0 ? '✓ 校验通过' : '⚠ 配置校验'}
        </span>
        <button className={`btn ${mode === 'raw' ? 'btn-primary' : ''}`} style={{ fontSize: 10 }}
          onClick={() => switchMode('raw')}>配置表检查</button>
        <button className={`btn ${mode === 'processed' ? 'btn-primary' : ''}`} style={{ fontSize: 10 }}
          onClick={() => switchMode('processed')}>导出数据检查</button>
        {data && <span style={{ fontSize: 11, color: 'var(--text-dim)', fontFamily: 'var(--font-mono)' }}>{data.total_errors} 个问题 · {data.files_affected} 个文件</span>}
        <span className="spacer" />
        <button className="btn" onClick={() => runCheck()} style={{ fontSize: 10 }}>重新校验</button>
        <button className="btn" onClick={() => setActivePanel(null)} style={{ fontSize: 10 }}>关闭</button>
      </div>

      {loading && <div className="loading"><div className="spinner" /> 扫描中...</div>}

      {data && data.total_errors === 0 && (
        <div className="empty-state" style={{ padding: 24 }}>
          <div style={{ fontSize: 24, opacity: 0.3, marginBottom: 8 }}>✓</div>
          <div>{mode === 'processed' ? '导出数据无悬空引用' : '导表检查通过，未发现问题'}</div>
        </div>
      )}

      {data && data.error && (
        <div style={{ padding: '8px 16px', color: 'var(--error)', fontSize: 12 }}>{data.error}</div>
      )}

      {data && data.skipped && data.skipped.length > 0 && (
        <div style={{ padding: '4px 16px', fontSize: 10, color: 'var(--text-dim)' }}>
          跳过: {data.skipped.map(s => `${s.rule}(${s.reason})`).join('、')}
        </div>
      )}

      {data && data.files && data.files.map(f => (
        <details key={f.file} style={{ borderBottom: '1px solid rgba(255,255,255,0.03)' }}>
          <summary className="val-file-header" style={{ outline: 'none' }}>
            <span style={{ color: 'var(--error)', fontWeight: 600 }}>{f.count}</span>
            <span>{f.file}</span>
            <span style={{ color: 'var(--text-dim)', fontSize: 10 }}>
              {Object.entries(f.by_type || {}).map(([k, v]) => `${k}:${v}`).join(', ')}
            </span>
          </summary>
          {Object.entries(f.by_sheet || {}).map(([sn, errs]) => (
            <div key={sn} style={{ marginLeft: 12, padding: '2px 0' }}>
              {sn && <div style={{ padding: '3px 16px', fontSize: 10, color: 'var(--text-dim)', fontFamily: 'var(--font-mono)' }}>{sn} ({errs.length})</div>}
              <table className="val-error-table">
                <thead><tr><th>错误</th><th>位置</th><th>行</th><th>列</th><th>字段</th><th>详情</th></tr></thead>
                <tbody>
                  {errs.map((e, i) => (
                    <tr key={i}>
                      <td className="err-type" style={e.level === 'warning' ? { color: 'var(--warn)' } : undefined}>{e.error}</td>
                      <td style={{ color: 'var(--warn)' }}>{e.col_ref}</td>
                      <td>{e.row}</td>
                      <td>{e.col}</td>
                      <td style={{ color: 'var(--text-dim)' }}>{e.col_header || '—'}</td>
                      <td className="err-detail" title={e.message || ''}>{e.message || '—'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ))}
        </details>
      ))}
    </div>
  )
}
