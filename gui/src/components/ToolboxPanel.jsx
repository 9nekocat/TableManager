import { useState } from 'react'
import { useApp } from '../context/AppContext'

const SECTIONS = [
  { key: 'recalc', label: '🔧 Effect #N/A 修复重算' },
  { key: 'check', label: '🔍 查 #N/A' },
]

function ErrorTable({ counts, details }) {
  if (!counts || Object.keys(counts).length === 0) {
    return <div className="empty-state" style={{ padding: 16 }}><div>✓ 未发现公式错误缓存</div></div>
  }
  return (
    <div style={{ padding: '0 16px 12px' }}>
      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
        <thead>
          <tr style={{ borderBottom: '1px solid var(--border)', textAlign: 'left' }}>
            <th style={{ padding: '6px 8px' }}>Sheet</th>
            <th style={{ padding: '6px 8px' }}>错误数</th>
          </tr>
        </thead>
        <tbody>
          {Object.entries(counts).map(([sheet, n]) => (
            <tr key={sheet} style={{ borderBottom: '1px solid var(--border)' }}>
              <td style={{ padding: '6px 8px', fontFamily: 'var(--font-mono)' }}>{sheet}</td>
              <td style={{ padding: '6px 8px', color: 'var(--error)' }}>{n}</td>
            </tr>
          ))}
        </tbody>
      </table>
      {details && details.length > 0 && (
        <>
          <div style={{ fontSize: 11, color: 'var(--text-dim)', margin: '8px 0 4px' }}>
            明细（前 {details.length} 个）：
          </div>
          <div style={{ maxHeight: 260, overflowY: 'auto', fontSize: 11, fontFamily: 'var(--font-mono)' }}>
            {details.map((d, i) => (
              <div key={i} style={{ padding: '2px 4px', color: 'var(--error)' }}>
                {d.sheet}!{d.cell} = {d.error}
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  )
}

function RecalcCard({ apiFetch }) {
  const [status, setStatus] = useState(null)
  const [result, setResult] = useState(null)
  const [loading, setLoading] = useState(false)

  const refreshStatus = async () => {
    const res = await apiFetch('/toolbox/status')
    setStatus(res)
  }

  const run = async () => {
    setLoading(true)
    try {
      const res = await apiFetch('/toolbox/effect/recalc', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({}),
      }, 300000)
      setResult(res)
      refreshStatus()
    } catch (e) {
      setResult({ ok: false, error: e.message })
    } finally {
      setLoading(false)
    }
  }

  if (!status) refreshStatus()

  return (
    <div className="merge-card">
      <div style={{ padding: '10px 16px', fontSize: 12, color: 'var(--text-dim)', lineHeight: 1.6 }}>
        针对 <span style={{ fontFamily: 'var(--font-mono)' }}>Effect.xlsx</span>（#Main / #事件效果 / #BOSS）的
        Params / ConditionIDs 公式列（TEXTJOIN+XLOOKUP）。修复重算会：
        <ol style={{ margin: '6px 0 0 16px', padding: 0 }}>
          <li>把 <span style={{ fontFamily: 'var(--font-mono)' }}>IF(X="无",0,…)</span> 升级为 <span style={{ fontFamily: 'var(--font-mono)' }}>IF(OR(X="无",X=""),0,…)</span>（名称列留空不再触发 XLOOKUP(0) → #N/A）</li>
          <li>COM 读回写 + 全量重算 + Excel 原生保存（备份在 Raw 目录外）</li>
        </ol>
        <div style={{ marginTop: 8, padding: '8px 10px', background: 'rgba(255, 180, 60, 0.12)', border: '1px solid rgba(255, 180, 60, 0.4)', borderRadius: 6 }}>
          ⚠ 本工具已为全部名称列 XLOOKUP 加上空值保护（<span style={{ fontFamily: 'var(--font-mono)' }}>IF(OR(X="无",X=""),0,…)</span>）。
          若在 Excel 中全量重算（F9 / Ctrl+Alt+F9）后出现大量 #N/A，请<b>勿保存</b>，
          关闭重开或运行本工具修复即可恢复。
        </div>
      </div>

      {status && (
        <div style={{ padding: '0 16px 8px', fontSize: 12 }}>
          {status.open_in_excel && (
            <div style={{ color: 'var(--error)', marginBottom: 6 }}>⚠ 文件正被 Excel 打开，需先关闭才能执行重算</div>
          )}
          <div>
            当前错误缓存：{status.total_errors} 个
            {status.error_counts && Object.entries(status.error_counts).map(([s, n]) => (
              <span key={s} style={{ marginLeft: 8, color: n ? 'var(--error)' : 'var(--success)', fontFamily: 'var(--font-mono)' }}>
                {s}: {n}
              </span>
            ))}
          </div>
        </div>
      )}

      {loading && <div className="loading"><div className="spinner" /> 重算中（打开 Excel 计算，约 10~30 秒）...</div>}

      {result && !loading && (
        <div style={{ padding: '0 16px 12px', fontSize: 12 }}>
          {result.ok ? (
            <>
              <div style={{ marginBottom: 4 }}>
                重算完成：<b style={{ color: 'var(--error)' }}>{Object.values(result.before || {}).reduce((a, b) => a + b, 0)}</b>
                {' → '}<b style={{ color: 'var(--success)' }}>{Object.values(result.after || {}).reduce((a, b) => a + b, 0)}</b> 个错误缓存
                <span style={{ marginLeft: 8, color: 'var(--text-dim)' }}>OR 保护 {result.or_protections} 处</span>
              </div>
              {result.backup_path && (
                <div style={{ color: 'var(--text-dim)', marginBottom: 4 }}>备份：{result.backup_path}</div>
              )}
              {result.residual_count > 0 ? (
                <>
                  <div style={{ color: 'var(--error)', margin: '6px 0' }}>{result.note}</div>
                  <ErrorTable counts={result.after} details={result.residual_details} />
                </>
              ) : (
                <div style={{ color: 'var(--success)' }}>{result.note}</div>
              )}
            </>
          ) : (
            <div style={{ color: 'var(--error)' }}>执行失败：{result.error}</div>
          )}
        </div>
      )}

      <div style={{ padding: '0 16px 14px' }}>
        <button className="btn btn-primary" onClick={run} disabled={loading || (status && status.open_in_excel)}>
          ▶ 执行修复重算
        </button>
      </div>
    </div>
  )
}

function CheckNaCard({ apiFetch, tables }) {
  const [name, setName] = useState('Effect.xlsx')
  const [result, setResult] = useState(null)
  const [loading, setLoading] = useState(false)

  const run = async () => {
    if (!name) return
    setLoading(true)
    try {
      const res = await apiFetch('/toolbox/check-na', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name }),
      }, 120000)
      setResult(res)
    } catch (e) {
      setResult({ ok: false, error: e.message })
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="merge-card">
      <div style={{ padding: '10px 16px', fontSize: 12, color: 'var(--text-dim)', lineHeight: 1.6 }}>
        对任意 Raw 配置表扫描公式错误缓存（#N/A / #REF! / #VALUE! 等），无需打开 Excel。
        适用于 <span style={{ fontFamily: 'var(--font-mono)' }}>Effect.xlsx</span>、
        <span style={{ fontFamily: 'var(--font-mono)' }}>Card.xlsx</span>、<span style={{ fontFamily: 'var(--font-mono)' }}>Event.xlsx</span> 等含公式的表。
      </div>

      <div style={{ padding: '0 16px 10px', display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
        <select
          value={name}
          onChange={e => { setName(e.target.value); setResult(null) }}
          style={{ background: 'var(--bg)', color: 'var(--text)', border: '1px solid var(--border)', borderRadius: 'var(--radius)', padding: '5px 8px', fontSize: 12, maxWidth: 260 }}
        >
          {(tables && tables.length ? tables : [{ file: 'Effect.xlsx' }]).map(t => (
            <option key={t.file} value={t.file}>{t.file}</option>
          ))}
        </select>
        <button className="btn btn-primary" onClick={run} disabled={loading || !name}>扫描 #N/A</button>
        {loading && <span className="spinner" style={{ width: 14, height: 14 }} />}
      </div>

      {result && !loading && (
        <div style={{ padding: '0 16px 14px', fontSize: 12 }}>
          {result.ok ? (
            <>
              <div style={{ marginBottom: 6 }}>
                共 <b style={{ color: result.total_errors ? 'var(--error)' : 'var(--success)' }}>{result.total_errors}</b> 个公式错误缓存
              </div>
              <ErrorTable counts={result.error_counts} details={result.details} />
            </>
          ) : (
            <div style={{ color: 'var(--error)' }}>{result.error}</div>
          )}
        </div>
      )}
    </div>
  )
}

export default function ToolboxPanel() {
  const { apiFetch, setActivePanel, flatTableData } = useApp()
  const [section, setSection] = useState('recalc')

  return (
    <div className="panel">
      <div className="panel-header">
        <span className="val-title" style={{ color: 'var(--text)' }}>🛠 妙妙工具箱</span>
        {SECTIONS.map(s => (
          <button key={s.key} className={`btn ${section === s.key ? 'btn-primary' : ''}`} style={{ fontSize: 10 }}
            onClick={() => setSection(s.key)}>
            {s.label}
          </button>
        ))}
        <span className="spacer" />
        <button className="btn" onClick={() => setActivePanel(null)} style={{ fontSize: 10 }}>关闭</button>
      </div>

      {section === 'recalc' && <RecalcCard apiFetch={apiFetch} />}
      {section === 'check' && <CheckNaCard apiFetch={apiFetch} tables={flatTableData} />}
    </div>
  )
}
