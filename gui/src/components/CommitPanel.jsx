import { useState, useEffect } from 'react'
import { useApp } from '../context/AppContext'

export default function CommitPanel() {
  const { svnStatus, updateInfo, apiFetch, showToast, refreshSvnStatus, setActivePanel, runSvnUpdate } = useApp()
  const [updating, setUpdating] = useState(false)
  const [message, setMessage] = useState('')
  const [checkedFiles, setCheckedFiles] = useState({})
  const [submitting, setSubmitting] = useState(false)

  const allItems = svnStatus?.all || []
  const changeFiles = allItems.filter(i => ['M', 'A', 'D'].includes(i.status))
  const conflicted = allItems.filter(i => i.status === 'C')

  useEffect(() => {
    const initial = {}
    changeFiles.forEach(f => { initial[f.path] = true })
    setCheckedFiles(initial)
  }, [svnStatus])

  const checked = Object.values(checkedFiles).filter(Boolean).length
  const canSubmit = checked > 0 && message.trim() && !submitting

  const doCommit = async () => {
    const files = Object.entries(checkedFiles).filter(([, v]) => v).map(([k]) => k)
    if (!files.length || !message.trim()) return
    setSubmitting(true)
    try {
      const data = await apiFetch('/svn/commit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message: message.trim(), files }),
      }, 120000)
      if (data.ok) {
        showToast(`提交成功 r${data.revision}`, 'success')
        setActivePanel(null)
        refreshSvnStatus()
      } else {
        showToast('提交失败: ' + data.error, 'error')
      }
    } catch (e) {
      showToast('提交失败: ' + e.message, 'error')
    } finally {
      setSubmitting(false)
    }
  }

  return (
    <div className="panel">
      <div className="panel-header">
        <span>SVN — {changeFiles.length} 个文件</span>
        {updateInfo?.has_update && <span style={{ color: 'var(--warn)', fontSize: 11 }}>远端 r{updateInfo.remote_rev} (落后 {updateInfo.behind_by})</span>}
        <span className="spacer" />
        <button className="btn btn-primary" onClick={async () => { setUpdating(true); await runSvnUpdate(); setUpdating(false) }}
          disabled={updating} style={{ fontSize: 10 }}>
          {updating ? '更新中...' : 'SVN Update'}
        </button>
        <button className="btn" onClick={() => setActivePanel(null)} style={{ fontSize: 10 }}>关闭</button>
      </div>

      {updateInfo?.has_update && (
        <div className="commit-warning">
          &#x26A0; 本地 r{updateInfo.local_rev} 落后远端 r{updateInfo.remote_rev} (落后 {updateInfo.behind_by} 个版本)
        </div>
      )}

      {conflicted.map(f => (
        <div className="commit-file-row" key={f.path}>
          <input type="checkbox" disabled />
          <span className="file-tag tag-C" style={{ background: 'var(--error-bg)', color: 'var(--error)' }}>C</span>
          <span className="file-name">{f.path}</span>
        </div>
      ))}

      {changeFiles.map(f => (
        <div className="commit-file-row" key={f.path}>
          <input type="checkbox" checked={!!checkedFiles[f.path]}
            onChange={e => setCheckedFiles(prev => ({ ...prev, [f.path]: e.target.checked }))} />
          <span className={`file-tag tag-${f.status}`}>{f.status}</span>
          <span className="file-name">{f.path}</span>
        </div>
      ))}

      {changeFiles.length === 0 && !conflicted.length && (
        <div className="empty-state" style={{ padding: 20 }}>无变更文件</div>
      )}

      <div className="commit-msg-row">
        <input type="text" placeholder="输入 commit message..." value={message}
          onChange={e => setMessage(e.target.value)} />
        <button className={`btn ${canSubmit ? 'btn-primary' : ''}`}
          disabled={!canSubmit} onClick={doCommit}>
          {submitting ? '提交中...' : '一键提交'}
        </button>
      </div>
    </div>
  )
}
