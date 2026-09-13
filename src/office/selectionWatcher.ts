import type { DiagramPayload } from '../metadata/payload'

export type DiagramSelectionWatcher = (() => void) & { refresh: () => void }

export type SelectionOrigin =
  | 'initial'
  | 'refresh'
  | 'office-event'
  | 'pane-focus'
  | 'document-poll'
  | 'background-poll'

// Rejected snapshots must not become the deduplication baseline.
export type SelectionReceiver = (payload: DiagramPayload | null, origin: SelectionOrigin) => boolean | void

export interface SelectionWatchOptions {
  isPaused?: () => boolean
  onSelectionChange?: (origin: SelectionOrigin) => void
}

export function isDocumentSelectionGesture(origin: SelectionOrigin): boolean {
  return origin === 'office-event' || origin === 'pane-focus' || origin === 'document-poll'
}

function mergeOrigin(previous: SelectionOrigin | undefined, next: SelectionOrigin): SelectionOrigin {
  if (!previous) return next
  if (isDocumentSelectionGesture(next)) return next
  if (isDocumentSelectionGesture(previous)) return previous
  if (previous === 'refresh' || next === 'refresh') return 'refresh'
  return next
}

export function watchDiagramSelection(
  getSelectedDiagram: () => Promise<DiagramPayload | null>,
  onSelected: SelectionReceiver,
  onError: (error: Error) => void,
  options: SelectionWatchOptions & { pollIntervalMs?: number } = {},
): DiagramSelectionWatcher {
  if (typeof Office === 'undefined' || !Office.context?.document) {
    return Object.assign(() => undefined, { refresh: () => undefined })
  }

  let active = true
  let checking = false
  let queuedOrigin: SelectionOrigin | undefined
  let queuedDeliverUnchanged = false
  let lastSelectionId: string | null | undefined
  let readFailed = false

  const checkSelection = async (origin: SelectionOrigin) => {
    if (!active) return
    queuedOrigin = mergeOrigin(queuedOrigin, origin)
    queuedDeliverUnchanged ||= origin !== 'background-poll' && origin !== 'document-poll'
    if (checking || options.isPaused?.()) return

    checking = true
    const readOrigin = queuedOrigin
    const deliverUnchanged = queuedDeliverUnchanged
    queuedOrigin = undefined
    queuedDeliverUnchanged = false
    try {
      const payload = await getSelectedDiagram()
      if (options.isPaused?.()) queuedOrigin = mergeOrigin(readOrigin, queuedOrigin ?? readOrigin)
      if (active && !queuedOrigin) {
        const id = payload?.id ?? null
        const changed = id !== lastSelectionId
        if (changed || deliverUnchanged || readFailed) {
          const accepted = onSelected(payload, readOrigin) !== false
          if (accepted) {
            lastSelectionId = id
            readFailed = false
          } else {
            // A write may have changed the editor's target since our last read.
            // Re-establish agreement rather than suppressing the next gesture.
            lastSelectionId = undefined
          }
        }
      }
    } catch (error) {
      if (options.isPaused?.()) queuedOrigin = mergeOrigin(readOrigin, queuedOrigin ?? readOrigin)
      if (active && !queuedOrigin) {
        readFailed = true
        onError(error instanceof Error ? error : new Error('Unable to read selected diagram.'))
      }
    } finally {
      checking = false
      if (queuedOrigin) {
        queuedOrigin = mergeOrigin(readOrigin, queuedOrigin)
        queuedDeliverUnchanged ||= deliverUnchanged
      }
      if (active && queuedOrigin && !options.isPaused?.()) {
        void checkSelection(queuedOrigin)
      }
    }
  }

  const checkOnFocus = () => {
    if (!options.isPaused?.() && document.visibilityState !== 'hidden') {
      void checkSelection('pane-focus')
    }
  }
  // Excel can change the active shape without changing the selected cell range.
  // WebView focus is not a reliable signal of workbook interaction in desktop
  // Excel. Poll visible panes, but don't treat replacement-driven nulls as clicks.
  const pollTimer = options.pollIntervalMs === undefined ? undefined : window.setInterval(() => {
    if (!checking && !options.isPaused?.() && document.visibilityState !== 'hidden') {
      void checkSelection(document.hasFocus() ? 'background-poll' : 'document-poll')
    }
  }, options.pollIntervalMs)
  if (pollTimer !== undefined) window.addEventListener('focus', checkOnFocus)

  const handler = () => {
    if (!active) return
    options.onSelectionChange?.('office-event')
    void checkSelection('office-event')
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
  void checkSelection('initial')

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
  return Object.assign(stop, { refresh: () => { void checkSelection('refresh') } })
}
