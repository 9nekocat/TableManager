import { useEffect, useState } from 'react'
import { useApp } from '../context/AppContext'

export default function MergePanel() {
  const {
    mergeStatus, apiFetch, showToast, refreshTables, refreshSvnStatus,
    refreshMergeStatus, setActivePanel,
  } = useApp()
  const [loading, setLoading] = useState({})
  const [lastResult, setLastResult] = useState(null)
  const [now, setNow] = useState(Date.now())
  const [conflictDefault, setConflictDefault] = useState('theirs')

  const busyText = {
    recover: 'Recovering changes from merge preview and recalculating...',
    preview: '正在生成并打开预览文件...',
    resolve: '正在写回合并结果、刷新公式缓存并清理冲突文件...',
  }

  useEffect(() => {
    if (!Object.values(loading).some(Boolean)) return undefined
    const timer = setInterval(() => setNow(Date.now()), 1000)
    return () => clearInterval(timer)
  }, [loading])

  const setBusy = (file, phase) => {
    setLoading(prev => ({ ...prev, [file]: phase ? { phase, startedAt: Date.now() } : null }))
  }

  const refreshAfterMergeAction = async () => {
    await Promise.all([
      refreshTables(),
      refreshSvnStatus(),
      refreshMergeStatus(),
    ])
  }

  const elapsedText = (state) => {
    if (!state?.startedAt) return ''
    const seconds = Math.max(0, Math.floor((now - state.startedAt) / 1000))
    if (seconds < 60) return `${seconds}s`
    return `${Math.floor(seconds / 60)}m ${String(seconds % 60).padStart(2, '0')}s`
  }

  const previewMerge = async (file) => {
    setLastResult(null)
    setBusy(file, 'preview')
    try {
      const data = await apiFetch(`/merge/preview/${encodeURIComponent(file)}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ conflict_default: conflictDefault }),
      }, 60000)
      if (data.error) { showToast(data.error, 'error') }
      else {
        showToast('合并预览已打开，确认 [最终] 行后点击“写回并重算”', 'success')
        const label = conflictDefault === 'mine' ? 'mine' : (conflictDefault === 'empty' ? '留空' : 'theirs')
        setLastResult({ type: 'success', message: `${file} 的合并预览已生成并打开，冲突默认采用 ${label}` })
        await refreshMergeStatus()
      }
    } catch (e) {
      setLastResult({ type: 'error', message: `${file} 预览生成失败：${e.message}` })
      showToast('生成失败: ' + e.message, 'error')
    } finally {
      setBusy(file, null)
    }
  }

  const resolveMerge = async (file) => {
    setLastResult(null)
    setBusy(file, 'resolve')
    try {
      const data = await apiFetch(`/merge/resolve/${encodeURIComponent(file)}`, { method: 'POST' }, 180000)
      if (data.error) { showToast(data.error, 'error') }
      else {
        const cache = data.formula_cache
        if (cache && cache.ok === false) {
          const detail = cache.missing_important_caches ? `，缺失 ${cache.missing_important_caches} 个 RowName 缓存` : ''
          const message = `合并完成，但公式缓存刷新失败${detail}：请用 Excel 打开保存后再比对`
          setLastResult({ type: 'error', message: `${file}：${message}` })
          showToast(message, 'error')
        } else {
          setLastResult({ type: 'success', message: `${file} 写回完成，公式缓存已刷新，冲突状态已重新扫描` })
          showToast('合并完成，公式缓存已刷新', 'success')
        }
      }
    } catch (e) {
      setLastResult({
        type: 'error',
        message: `${file} 写回失败或请求超时：${e.message}。已重新扫描状态；如果 Excel 仍在后台重算，请稍后点“刷新”。`,
      })
      showToast('合并失败: ' + e.message, 'error')
    } finally {
      await refreshAfterMergeAction()
      setBusy(file, null)
    }
  }

  const recoverMerge = async (file) => {
    setLastResult(null)
    setBusy(file, 'recover')
    try {
      const data = await apiFetch(`/merge/recover/${encodeURIComponent(file)}`, { method: 'POST' }, 180000)
      const backup = data.backup_path ? `, backup: ${data.backup_path}` : ''
      setLastResult({ type: 'success', message: `${file} recovered from merge preview${backup}` })
      showToast('Recovered from merge preview', 'success')
    } catch (e) {
      setLastResult({ type: 'error', message: `${file} recover failed: ${e.message}` })
      showToast('Recover failed: ' + e.message, 'error')
    } finally {
      await refreshAfterMergeAction()
      setBusy(file, null)
    }
  }

  const openPreview = async (previewName) => {
    try {
      await apiFetch(`/open/${encodeURIComponent(previewName)}`, { method: 'POST' }, 15000)
    } catch (e) {
      showToast('Open preview failed: ' + e.message, 'error')
    }
  }

  const abortMerge = async (file) => {
    if (!window.confirm(`放弃合并并清理冲突文件？\n${file}`)) return
    try {
      await apiFetch(`/merge/abort/${encodeURIComponent(file)}`, { method: 'POST' }, 15000)
      setLastResult({ type: 'success', message: `${file} 的合并预览已放弃` })
      await refreshMergeStatus()
    } catch (e) {
      setLastResult({ type: 'error', message: `${file} 放弃合并失败：${e.message}` })
      showToast('操作失败: ' + e.message, 'error')
    }
  }

  const total = mergeStatus.total || mergeStatus.conflicts?.length || 0
  const recoverablePreviews = mergeStatus.recoverable_previews || []

  return (
    <div className="panel">
      <div className="panel-header">
        <span style={{ color: total ? 'var(--error)' : 'var(--success)' }}>
          {total ? `冲突文件 (${total})` : '✓ 无冲突'}
        </span>
        <span className="spacer" />
        <button className="btn" onClick={() => setActivePanel(null)} style={{ fontSize: 10 }}>关闭</button>
      </div>

      {lastResult && (
        <div className={`merge-result ${lastResult.type}`}>
          {lastResult.message}
        </div>
      )}

      {total === 0 && <div className="empty-state" style={{ padding: 20 }}>所有文件已同步</div>}

      {(mergeStatus.conflicts || []).map(c => (
        <div key={c.file} className={`merge-card${loading[c.file] ? ' is-busy' : ''}`}>
          {(() => {
            const state = loading[c.file]
            const phase = state?.phase
            const elapsed = elapsedText(state)
            return (
              <>
          <div style={{ fontWeight: 600, marginBottom: 3 }}>{c.file}</div>
          <div style={{ fontSize: 10, color: 'var(--text-dim)' }}>
            mine: {c.mine_path}<br />theirs: {c.theirs_path}<br />base: {c.base_path || '(none)'}
          </div>
          <div className={`merge-state ${c.preview_exists ? 'ready' : ''}`}>
            {phase ? `${busyText[phase]} 已用 ${elapsed || '0s'}${phase === 'resolve' ? '，Excel 重算可能需要较久，请勿重复点击' : ''}` : (c.preview_exists ? '预览已生成，可直接应用' : '尚未生成预览')}
          </div>
          <label className="merge-option">
            <span>冲突默认</span>
            <select value={conflictDefault} onChange={e => setConflictDefault(e.target.value)}
              disabled={!!loading[c.file]}>
              <option value="theirs">采用 theirs</option>
              <option value="mine">采用 mine</option>
              <option value="empty">默认留空</option>
            </select>
          </label>
          <div className="merge-actions">
            <button className="btn btn-primary" onClick={() => previewMerge(c.file)}
              disabled={!!loading[c.file]} style={{ fontSize: 10 }}>
              {phase === 'preview' ? '生成中...' : (c.preview_exists ? '打开预览' : '生成预览')}
            </button>
            <button className="btn" onClick={() => resolveMerge(c.file)}
              disabled={!!loading[c.file] || !c.preview_exists} style={{ fontSize: 10 }}
              title={c.preview_exists ? '写回合并结果并刷新公式缓存' : '请先生成合并预览'}>
              {phase === 'resolve' ? `写回中 ${elapsed || ''}` : '写回并重算'}
            </button>
            <button className="btn" onClick={() => recoverMerge(c.file)}
              disabled={!!loading[c.file] || !c.preview_exists} style={{ fontSize: 10 }}
              title="Re-apply the saved *_merge.xlsx final rows">
              {phase === 'recover' ? `Recovering ${elapsed || ''}` : '从预览恢复'}
            </button>
            <button className="btn btn-danger" onClick={() => abortMerge(c.file)} style={{ fontSize: 10 }}>
              放弃
            </button>
          </div>
              </>
            )
          })()}
        </div>
      ))}

      {recoverablePreviews.length > 0 && (
        <div className="merge-recovery-section">
          <div className="merge-recovery-title">Recoverable merge previews</div>
          {recoverablePreviews.map(c => {
            const cardKey = c.preview_name || c.file
            const state = loading[cardKey]
            const phase = state?.phase
            const elapsed = elapsedText(state)
            return (
              <div key={cardKey} className={`merge-card${state ? ' is-busy' : ''}`}>
                <div style={{ fontWeight: 600, marginBottom: 3 }}>{c.file}</div>
                <div style={{ fontSize: 10, color: 'var(--text-dim)' }}>preview: {c.preview_path}</div>
                <div className="merge-state ready">
                  {phase ? `${busyText[phase]} elapsed ${elapsed || '0s'}` : 'Can re-apply final rows from this *_merge.xlsx'}
                </div>
                <div className="merge-actions">
                  <button className="btn" onClick={() => openPreview(cardKey)}
                    disabled={!!state} style={{ fontSize: 10 }}>
                    打开预览
                  </button>
                  <button className="btn btn-primary" onClick={() => recoverMerge(cardKey)}
                    disabled={!!state} style={{ fontSize: 10 }}>
                    {phase === 'recover' ? `恢复中 ${elapsed || ''}` : '恢复修改'}
                  </button>
                  <button className="btn btn-danger" onClick={() => abortMerge(cardKey)}
                    disabled={!!state} style={{ fontSize: 10 }}>
                    删除预览
                  </button>
                </div>
              </div>
            )
          })}
        </div>
      )}
    </div>
  )
}
