import { useEffect } from 'react'
import { AppProvider, useApp } from '../context/AppContext'
import { useWindowSize } from '../hooks/useWindowSize'
import ParticleBg from './ParticleBg'
import Sidebar from './Sidebar'
import Toolbar from './Toolbar'
import TableGrid from './TableGrid'
import CommitPanel from './CommitPanel'
import DiffPanel from './DiffPanel'
import ValidationPanel from './ValidationPanel'
import RefsPanel from './RefsPanel'
import MergePanel from './MergePanel'
import ToolboxPanel from './ToolboxPanel'
import Toast from './Toast'

function AppInner() {
  const { activePanel, refreshTables, refreshSvnStatus, checkUpdates, refreshMergeStatus, checkEnvironment } = useApp()

  useWindowSize()

  useEffect(() => {
    checkEnvironment()
    refreshTables()
    refreshSvnStatus()
    refreshMergeStatus()
    const t1 = setTimeout(checkUpdates, 2000)
    const t2 = setInterval(checkUpdates, 300000)
    return () => { clearTimeout(t1); clearInterval(t2) }
  }, [checkEnvironment, refreshTables, refreshSvnStatus, checkUpdates, refreshMergeStatus])

  return (
    <>
      <ParticleBg />
      <Sidebar />
      <main className="main">
        <Toolbar />
        <TableGrid />
        {activePanel === 'commit' && <CommitPanel />}
        {activePanel === 'diff' && <DiffPanel />}
        {activePanel === 'validation' && <ValidationPanel />}
        {activePanel === 'refs' && <RefsPanel />}
        {activePanel === 'merge' && <MergePanel />}
        {activePanel === 'toolbox' && <ToolboxPanel />}
      </main>
      <Toast />
    </>
  )
}

export default function App() {
  return (
    <AppProvider>
      <AppInner />
    </AppProvider>
  )
}
