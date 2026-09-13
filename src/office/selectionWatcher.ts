import type { DiagramPayload } from '../metadata/payload'

export type DiagramSelectionWatcher = (() => void) & { refresh: () => void }

export interface SelectionWatchOptions {
  isPaused?: () => boolean
  onSelectionChange?: (fromDocument?: boolean) => void
}

export function watchDiagramSelection(
  getSelectedDiagram: () => Promise<DiagramPayload | null>,
  onSelected: (payload: DiagramPayload | null) => void,
  onError: (error: Error) => void,
  options: SelectionWatchOptions & { pollIntervalMs?: number } = {},
): DiagramSelectionWatcher {
  if (typeof Office === 'undefined' || !Office.context?.document) {
    return Object.assign(() => undefined, { refresh: () => undefined })
  }

  let active = true
  let checking = false
  let queued = false
  let queuedFromDocument = false
  let queuedRefresh = false
  let lastSelectionId: string | null | undefined
  let readFailed = false

  const checkSelection = async (fromDocument = false) => {
    if (!active) return
    queuedFromDocument ||= fromDocument
    queuedRefresh ||= !fromDocument
    if (checking || options.isPaused?.()) {
      queued = true
      return
    }

    checking = true
    queued = false
    const notifyOnChange = queuedFromDocument
    const deliverUnchanged = queuedRefresh
    queuedFromDocument = false
    queuedRefresh = false
    try {
      const payload = await getSelectedDiagram()
      if (options.isPaused?.()) queued = true
      if (active && !queued) {
        const id = payload?.id ?? null
        const changed = id !== lastSelectionId
        // Background probes must not repeatedly lock the editor or discard a draft.
        if (notifyOnChange && changed) options.onSelectionChange?.(true)
        lastSelectionId = id
        if (deliverUnchanged || changed || readFailed) onSelected(payload)
        readFailed = false
      }
    } catch (error) {
      if (options.isPaused?.()) queued = true
      if (active && !queued) {
        readFailed = true
        onError(error instanceof Error ? error : new Error('Unable to read selected diagram.'))
      }
    } finally {
      checking = false
      if (queued) {
        queuedFromDocument ||= notifyOnChange
        queuedRefresh ||= deliverUnchanged
      }
      if (active && queued && !options.isPaused?.()) {
        void checkSelection(queuedFromDocument)
      }
    }
  }

  const checkOnFocus = () => {
    if (!options.isPaused?.() && document.visibilityState !== 'hidden') {
      void checkSelection(true)
    }
  }
  // Excel can change the active shape without changing the selected cell range.
  // Only Excel opts in; avoid Office roundtrips while typing or writing an image.
  const pollTimer = options.pollIntervalMs === undefined ? undefined : window.setInterval(() => {
    if (!checking && !options.isPaused?.() && !document.hasFocus() && document.visibilityState !== 'hidden') {
      void checkSelection(true)
    }
  }, options.pollIntervalMs)
  if (pollTimer !== undefined) window.addEventListener('focus', checkOnFocus)

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
    if (pollTimer !== undefined) {
      window.clearInterval(pollTimer)
      window.removeEventListener('focus', checkOnFocus)
    }
    Office.context.document.removeHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      { handler },
    )
  }
  return Object.assign(stop, { refresh: () => { void checkSelection() } })
}
