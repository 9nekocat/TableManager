import { createContext, useContext, useState, useCallback, useEffect, useRef } from 'react'

const AppContext = createContext(null)

export function AppProvider({ children }) {
  const [categories, setCategories] = useState([])
  const [flatTableData, setFlatTableData] = useState([])
  const [svnStatus, setSvnStatus] = useState(null)
  const [mergeStatus, setMergeStatus] = useState({ total: 0, conflicts: [] })
  const [updateInfo, setUpdateInfo] = useState(null)
  const [activeCategory, setActiveCategory] = useState(null)
  const [narrowMode, setNarrowMode] = useState(window.innerWidth < 700)
  const [activePanel, setActivePanel] = useState(null)
  const [searchQuery, setSearchQuery] = useState('')
  const [toasts, setToasts] = useState([])
  const [theme, setTheme] = useState(() => {
    const saved = localStorage.getItem('tm-theme')
    if (saved === 'light' || saved === 'dark') return saved
    return window.matchMedia?.('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
  })
  const toastId = useRef(0)

  const API = '/api'

  useEffect(() => {
    document.documentElement.dataset.theme = theme
    localStorage.setItem('tm-theme', theme)
  }, [theme])

  const toggleTheme = useCallback(() => {
    setTheme(prev => prev === 'dark' ? 'light' : 'dark')
  }, [])

  const showToast = useCallback((message, type = 'success') => {
    const id = ++toastId.current
    setToasts(prev => [...prev, { id, message, type }])
    setTimeout(() => setToasts(prev => prev.filter(t => t.id !== id)), type === 'error' ? 6000 : 3000)
  }, [])

  const apiFetch = useCallback(async (url, options = {}, timeoutMs = 30000) => {
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), timeoutMs)
    try {
      const res = await fetch(`${API}${url}`, { ...options, signal: controller.signal })
      const text = await res.text()
      let data = {}
      if (text) {
        try {
          data = JSON.parse(text)
        } catch (_) {
          data = { error: text }
        }
      }
      if (!res.ok) {
        throw new Error(data.error || `请求失败 (${res.status})`)
      }
      return data
    } catch (e) {
      if (e.name === 'AbortError') {
        throw new Error('请求超时，请稍后重试或检查 Excel/SVN 是否卡住')
      }
      throw e
    } finally {
      clearTimeout(timer)
    }
  }, [API])

  const refreshTables = useCallback(async () => {
    try {
      const data = await apiFetch('/tables')
      setCategories(data)
      const flat = data.flatMap(c => c.tables.map(t => ({ ...t, catKey: c.key, catDisplay: c.display_name })))
      setFlatTableData(flat)
    } catch (e) {
      showToast('加载表格失败: ' + e.message, 'error')
    }
  }, [apiFetch, showToast])

  const checkEnvironment = useCallback(async () => {
    try {
      const data = await apiFetch('/env/status', {}, 300000)
      const svn = data.svn || {}
      if (svn.installed && svn.install_attempted) {
        showToast('已自动安装 SVN 命令行工具', 'success')
      } else if (!svn.installed && svn.error) {
        showToast('SVN 环境未就绪: ' + svn.error, 'error')
      }
      return data
    } catch (e) {
      showToast('环境检查失败: ' + e.message, 'error')
      return null
    }
  }, [apiFetch, showToast])

  const refreshSvnStatus = useCallback(async () => {
    try {
      const data = await apiFetch('/svn/status', {}, 15000)
      if (!data.error) setSvnStatus(data)
    } catch (_) { /* silent */ }
  }, [apiFetch])

  const checkUpdates = useCallback(async () => {
    try {
      const data = await apiFetch('/svn/check-updates', {}, 15000)
      if (!data.error) setUpdateInfo(data)
    } catch (_) { /* silent */ }
  }, [apiFetch])

  const refreshMergeStatus = useCallback(async () => {
    try {
      const data = await apiFetch('/merge/status', {}, 15000)
      if (!data.error) setMergeStatus(data)
    } catch (_) { /* silent */ }
  }, [apiFetch])

  const runSvnUpdate = useCallback(async () => {
    try {
      const data = await apiFetch('/svn/update', { method: 'POST' }, 120000)
      if (data.error) { showToast('更新失败: ' + data.error, 'error'); return }
      showToast(`更新完成 r${data.revision} (更新${data.updated} 冲突${data.conflicted})`, data.conflicted ? 'error' : 'success')
      await Promise.all([
        refreshTables(),
        refreshSvnStatus(),
        refreshMergeStatus(),
        checkUpdates(),
      ])
    } catch (e) {
      showToast('更新失败: ' + e.message, 'error')
    }
  }, [apiFetch, showToast, refreshTables, refreshSvnStatus, refreshMergeStatus, checkUpdates])

  return (
    <AppContext.Provider value={{
      categories, flatTableData, svnStatus, mergeStatus, updateInfo,
      activeCategory, narrowMode, activePanel, searchQuery, toasts, theme,
      showToast, apiFetch, setActiveCategory, setNarrowMode, setActivePanel,
      setSearchQuery,
      refreshTables, refreshSvnStatus, checkUpdates, refreshMergeStatus,
      runSvnUpdate, toggleTheme, checkEnvironment,
    }}>
      {children}
    </AppContext.Provider>
  )
}

export function useApp() {
  const ctx = useContext(AppContext)
  if (!ctx) throw new Error('useApp must be inside AppProvider')
  return ctx
}
