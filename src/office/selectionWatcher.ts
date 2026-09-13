import type { DiagramPayload } from '../metadata/payload'

export type DiagramSelectionWatcher = (() => void) & { refresh: () => void }

export function watchDiagramSelection(
  getSelectedDiagram: () => Promise<DiagramPayload | null>,
  onSelected: (payload: DiagramPayload | null) => void,
  onError: (error: Error) => void,
  options: { isPaused?: () => boolean; onSelectionChange?: () => void } = {},
): DiagramSelectionWatcher {
  if (typeof Office === 'undefined' || !Office.context?.document) {
    return Object.assign(() => undefined, { refresh: () => undefined })
  }

  let active = true
  let checking = false
  let queued = false

  const checkSelection = async () => {
    if (!active) return
    if (checking || options.isPaused?.()) {
      queued = true
      return
    }

    checking = true
    queued = false
    try {
      const payload = await getSelectedDiagram()
      if (options.isPaused?.()) queued = true
      if (active && !queued) {
        onSelected(payload)
      }
    } catch (error) {
      if (options.isPaused?.()) queued = true
      if (active && !queued) {
        onError(error instanceof Error ? error : new Error('Unable to read selected diagram.'))
      }
    } finally {
      checking = false
      if (active && queued && !options.isPaused?.()) {
        void checkSelection()
      }
    }
  }

  const handler = () => {
    if (!active) return
    options.onSelectionChange?.()
    void checkSelection()
  }

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    handler,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed && active) {
        onError(new Error(result.error.message))
      }
    },
  )
  void checkSelection()

  const stop = () => {
    active = false
    Office.context.document.removeHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      { handler },
    )
  }
  return Object.assign(stop, { refresh: () => { void checkSelection() } })
}
