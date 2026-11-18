import { useEffect } from 'react'
import { useApp } from '../context/AppContext'

export function useWindowSize() {
  const { setNarrowMode } = useApp()

  useEffect(() => {
    const check = () => setNarrowMode(window.innerWidth < 700)
    window.addEventListener('resize', check)
    return () => window.removeEventListener('resize', check)
  }, [setNarrowMode])
}
