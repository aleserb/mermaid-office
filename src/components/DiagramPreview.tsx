import { Button, MessageBar, MessageBarBody, Spinner } from '@fluentui/react-components'
import {
  ScaleFillRegular,
  ZoomInRegular,
  ZoomOutRegular,
} from '@fluentui/react-icons'
import { useEffect, useRef, useState, type PointerEvent } from 'react'
import './DiagramPreview.css'

const MIN_SCALE = 0.25
const MAX_SCALE = 4
const SCALE_STEP = 0.25

interface DiagramPreviewProps {
  svg: string
  loading: boolean
  error?: string
}

interface Point {
  x: number
  y: number
}

function clampScale(scale: number): number {
  return Math.min(MAX_SCALE, Math.max(MIN_SCALE, scale))
}

export function DiagramPreview({ svg, loading, error = '' }: DiagramPreviewProps) {
  const [scale, setScale] = useState(1)
  const [pan, setPan] = useState<Point>({ x: 0, y: 0 })
  const dragOrigin = useRef<Point | null>(null)
  const viewportRef = useRef<HTMLDivElement>(null)

  const zoom = (change: number) => {
    setScale((current) => clampScale(current + change))
  }

  const fit = () => {
    setScale(1)
    setPan({ x: 0, y: 0 })
  }

  const handlePointerDown = (event: PointerEvent<HTMLDivElement>) => {
    if (event.button !== 0 || scale <= 1) {
      return
    }
    event.currentTarget.setPointerCapture(event.pointerId)
    dragOrigin.current = {
      x: event.clientX - pan.x,
      y: event.clientY - pan.y,
    }
  }

  const handlePointerMove = (event: PointerEvent<HTMLDivElement>) => {
    if (!dragOrigin.current) {
      return
    }
    setPan({
      x: event.clientX - dragOrigin.current.x,
      y: event.clientY - dragOrigin.current.y,
    })
  }

  const stopDragging = (event: PointerEvent<HTMLDivElement>) => {
    dragOrigin.current = null
    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId)
    }
  }

  useEffect(() => {
    const viewport = viewportRef.current
    if (!viewport) {
      return
    }

    const handleWheel = (event: globalThis.WheelEvent) => {
      if (!event.ctrlKey && !event.metaKey) {
        return
      }
      event.preventDefault()
      setScale((current) =>
        clampScale(current + (event.deltaY < 0 ? SCALE_STEP : -SCALE_STEP)),
      )
    }

    viewport.addEventListener('wheel', handleWheel, { passive: false })
    return () => viewport.removeEventListener('wheel', handleWheel)
  }, [])

  return (
    <div className="diagram-preview">
      <div className="preview-toolbar" aria-label="Preview navigation">
        <Button
          appearance="subtle"
          aria-label="Zoom out"
          icon={<ZoomOutRegular />}
          disabled={scale === MIN_SCALE}
          onClick={() => zoom(-SCALE_STEP)}
        />
        <span className="preview-scale" aria-live="polite">
          {Math.round(scale * 100)}%
        </span>
        <Button
          appearance="subtle"
          aria-label="Zoom in"
          icon={<ZoomInRegular />}
          disabled={scale === MAX_SCALE}
          onClick={() => zoom(SCALE_STEP)}
        />
        <Button
          appearance="subtle"
          aria-label="Fit diagram"
          title="Fit diagram"
          icon={<ScaleFillRegular />}
          onClick={fit}
        />
      </div>

      <div
        ref={viewportRef}
        className={`preview-viewport${scale > 1 ? ' is-pannable' : ''}`}
        onDoubleClick={fit}
        onPointerDown={handlePointerDown}
        onPointerMove={handlePointerMove}
        onPointerUp={stopDragging}
        onPointerCancel={stopDragging}
      >
        {loading && !svg && <Spinner label="Rendering diagram" />}
        <div
          className="preview-stage"
          style={{ transform: `translate(${pan.x}px, ${pan.y}px) scale(${scale})` }}
        >
          {svg && (
            <div
              className="preview-svg"
              // Mermaid output is sanitized before display.
              dangerouslySetInnerHTML={{ __html: svg }}
            />
          )}
        </div>
      </div>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
    </div>
  )
}
