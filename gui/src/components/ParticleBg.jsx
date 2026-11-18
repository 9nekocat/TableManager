import { useEffect, useRef } from 'react'

const PARTICLE_COUNT = 96
const MAX_DIST = 135

export default function ParticleBg() {
  const canvasRef = useRef(null)

  useEffect(() => {
    const canvas = canvasRef.current
    if (!canvas) return
    const ctx = canvas.getContext('2d')
    let animId, particles = []
    let w, h

    function resize() {
      w = canvas.width = window.innerWidth
      h = canvas.height = window.innerHeight
    }
    resize()
    window.addEventListener('resize', resize)

    for (let i = 0; i < PARTICLE_COUNT; i++) {
      particles.push({
        x: Math.random() * w,
        y: Math.random() * h,
        vx: (Math.random() - 0.5) * 0.42,
        vy: (Math.random() - 0.5) * 0.22 - 0.12,
        r: Math.random() * 1.5 + 0.55,
        phase: Math.random() * Math.PI * 2,
        spin: (Math.random() - 0.5) * 0.018,
        kind: Math.random() > 0.78 ? 'spark' : 'petal',
      })
    }

    function frame() {
      const styles = getComputedStyle(document.documentElement)
      const particleColor = styles.getPropertyValue('--particle').trim() || 'rgba(255, 143, 189, 0.28)'
      const lineColor = styles.getPropertyValue('--particle-line').trim() || 'rgba(69, 182, 232, 0.12)'
      const isDark = document.documentElement.dataset.theme === 'dark'
      const t = performance.now() * 0.001

      ctx.clearRect(0, 0, w, h)
      for (let i = 0; i < particles.length; i++) {
        const p = particles[i]
        p.x += p.vx; p.y += p.vy
        p.phase += p.spin
        if (p.x < 0) p.x = w
        if (p.x > w) p.x = 0
        if (p.y < -10) p.y = h + 10
        if (p.y > h + 10) p.y = -10

        const pulse = 0.62 + Math.sin(t * 1.8 + p.phase) * 0.22
        ctx.beginPath()
        if (p.kind === 'spark') {
          const size = p.r * (isDark ? 2.1 : 1.7)
          ctx.moveTo(p.x, p.y - size)
          ctx.lineTo(p.x + size * 0.26, p.y - size * 0.26)
          ctx.lineTo(p.x + size, p.y)
          ctx.lineTo(p.x + size * 0.26, p.y + size * 0.26)
          ctx.lineTo(p.x, p.y + size)
          ctx.lineTo(p.x - size * 0.26, p.y + size * 0.26)
          ctx.lineTo(p.x - size, p.y)
          ctx.lineTo(p.x - size * 0.26, p.y - size * 0.26)
          ctx.closePath()
        } else if (isDark) {
          ctx.arc(p.x, p.y, p.r * 1.15, 0, Math.PI * 2)
        } else {
          ctx.ellipse(p.x, p.y, p.r * 2.15, p.r * 0.82, p.phase + (p.x + p.y) * 0.01, 0, Math.PI * 2)
        }
        ctx.fillStyle = particleColor
        ctx.globalAlpha = pulse
        ctx.fill()
        ctx.globalAlpha = 1

        if (isDark) {
          for (let j = i + 1; j < particles.length; j++) {
            const q = particles[j]
            const dx = p.x - q.x, dy = p.y - q.y
            const dist = Math.sqrt(dx * dx + dy * dy)
            if (dist < MAX_DIST) {
              ctx.beginPath()
              ctx.moveTo(p.x, p.y)
              ctx.lineTo(q.x, q.y)
              ctx.strokeStyle = lineColor
              ctx.globalAlpha = (1 - dist / MAX_DIST) * 0.82
              ctx.lineWidth = 0.4
              ctx.stroke()
              ctx.globalAlpha = 1
            }
          }
        } else if (p.kind === 'spark') {
          ctx.beginPath()
          ctx.arc(p.x, p.y, p.r * 4.4, 0, Math.PI * 2)
          ctx.strokeStyle = lineColor
          ctx.globalAlpha = 0.26 * pulse
          ctx.lineWidth = 0.5
          ctx.stroke()
          ctx.globalAlpha = 1
        }
      }
      animId = requestAnimationFrame(frame)
    }
    frame()

    return () => {
      cancelAnimationFrame(animId)
      window.removeEventListener('resize', resize)
    }
  }, [])

  return <canvas ref={canvasRef} id="particle-bg" />
}
