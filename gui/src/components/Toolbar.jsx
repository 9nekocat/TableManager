import { useApp } from '../context/AppContext'

export default function Toolbar() {
  const {
    setActivePanel, activePanel, updateInfo, svnStatus, mergeStatus,
    refreshTables, refreshSvnStatus, refreshMergeStatus, checkUpdates,
    theme, toggleTheme, searchQuery, setSearchQuery,
  } = useApp()

  const badgeCount = mergeStatus.total || 0
  const svnChanges = svnStatus ? svnStatus.total_changes || 0 : 0
  const behindBy = updateInfo?.behind_by || 0

  const doRefresh = async () => {
    await Promise.all([
      refreshTables(),
      refreshSvnStatus(),
      refreshMergeStatus(),
      checkUpdates(),
    ])
  }

  return (
    <div className="toolbar">
      <div className="search-box">
        <input type="text" placeholder="搜索表格名或 UE 表名..." value={searchQuery}
          onChange={e => setSearchQuery(e.target.value)} />
        {searchQuery && (
          <button className="search-clear" type="button" onClick={() => setSearchQuery('')} title="清空搜索" aria-label="清空搜索">
            &times;
          </button>
        )}
      </div>
      <span className="spacer" />

      <button className="btn" onClick={() => setActivePanel(activePanel === 'commit' ? null : 'commit')}
        style={{ position: 'relative' }} title="SVN 状态/提交">
        &#x1F4E6; SVN
        {behindBy > 0 && <span className="badge-update" style={{ display: 'inline' }}>{behindBy}</span>}
      </button>

      <button className="btn" onClick={doRefresh} title="刷新 SVN 状态">
        &#x1F504; 刷新
      </button>

      {badgeCount > 0 && (
        <button className="btn" onClick={() => setActivePanel(activePanel === 'merge' ? null : 'merge')} title="合并冲突">
          &#x2694; 冲突 <span style={{ background: 'var(--error)', color: '#fff', borderRadius: 8, padding: '0 5px', fontSize: 10 }}>{badgeCount}</span>
        </button>
      )}

      <button className={`btn ${activePanel === 'validation' ? 'btn-primary' : ''}`}
        onClick={() => setActivePanel(activePanel === 'validation' ? null : 'validation')}>
        &#x26A0; 校验全部
      </button>

      <button className={`btn ${activePanel === 'refs' ? 'btn-primary' : ''}`}
        onClick={() => setActivePanel(activePanel === 'refs' ? null : 'refs')}>
        &#x1F517; 引用
      </button>

      <button className={`btn ${activePanel === 'toolbox' ? 'btn-primary' : ''}`}
        onClick={() => setActivePanel(activePanel === 'toolbox' ? null : 'toolbox')}>
        &#x1F6E0; 妙妙工具箱
      </button>

      <button className="btn" onClick={() => setActivePanel(activePanel === 'diff' ? null : 'diff')}>
        &#x1F504; 版本比对
      </button>

      <button className="btn btn-icon theme-toggle" onClick={toggleTheme}
        title={theme === 'dark' ? '切换到亮色主题' : '切换到暗色主题'}
        aria-label={theme === 'dark' ? '切换到亮色主题' : '切换到暗色主题'}>
        {theme === 'dark' ? '☀︎' : '☾'}
      </button>
    </div>
  )
}
