import { useEffect } from 'react'
import { normalizeMermaidError } from '../mermaid/diagnostics'
import { renderMermaid } from '../mermaid/render'
import type { DiagramDraft } from '../office/diagramRequests'
import type { EditorAction } from './editorReducer'

export const LIVE_UPDATE_DELAY = 600

export function useDebouncedDiagramRender(
  draft: DiagramDraft,
  ready: boolean,
  dispatch: (action: EditorAction) => void,
): void {
  useEffect(() => {
    if (!ready) return
    let active = true
    const timer = window.setTimeout(async () => {
      try {
        const svg = await renderMermaid(draft.source, draft.theme, draft.settings)
        if (active) dispatch({ type: 'rendered', result: { draft, svg } })
      } catch (error) {
        if (active) dispatch({ type: 'render-failed', draft, diagnostic: normalizeMermaidError(error, draft.source) })
      }
    }, LIVE_UPDATE_DELAY)
    return () => {
      active = false
      window.clearTimeout(timer)
    }
  }, [draft, ready, dispatch])
}
