import {
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent,
  type PointerEvent,
  type ReactNode,
} from 'react'
import './SplitWorkspace.css'

interface SplitWorkspaceProps {
  ariaLabel: string
  className?: string
  left: ReactNode
  right: ReactNode
}

export function SplitWorkspace({
  ariaLabel,
  className = '',
  left,
  right,
}: SplitWorkspaceProps) {
  const [leftWidth, setLeftWidth] = useState(50)
  const containerRef = useRef<HTMLElement>(null)
  const dragging = useRef(false)

  const resizeFromPointer = (event: PointerEvent<HTMLDivElement>) => {
    if (!dragging.current || !containerRef.current) {
      return
    }
    const bounds = containerRef.current.getBoundingClientRect()
    const percentage = ((event.clientX - bounds.left) / bounds.width) * 100
    setLeftWidth(Math.min(75, Math.max(25, percentage)))
  }

  const stopResizing = (event: PointerEvent<HTMLDivElement>) => {
    dragging.current = false
    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId)
    }
  }

  const handleSeparatorKey = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') {
      return
    }
    event.preventDefault()
    setLeftWidth((current) =>
      Math.min(75, Math.max(25, current + (event.key === 'ArrowLeft' ? -5 : 5))),
    )
  }

  return (
    <section
      ref={containerRef}
      className={`split-workspace ${className}`.trim()}
      aria-label={ariaLabel}
      style={{ '--editor-pane-width': `${leftWidth}%` } as CSSProperties}
    >
      {left}
      <div
        className="split-separator"
        role="separator"
        aria-label="Resize editor and preview"
        aria-orientation="vertical"
        aria-valuemin={25}
        aria-valuemax={75}
        aria-valuenow={Math.round(leftWidth)}
        tabIndex={0}
        onKeyDown={handleSeparatorKey}
        onPointerDown={(event) => {
          dragging.current = true
          event.currentTarget.setPointerCapture(event.pointerId)
        }}
        onPointerMove={resizeFromPointer}
        onPointerUp={stopResizing}
        onPointerCancel={stopResizing}
      />
      {right}
    </section>
  )
}
