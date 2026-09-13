import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { createDiagramPayload, type DiagramPayload } from '../metadata/payload'
import { watchDiagramSelection, type DiagramSelectionWatcher, type SelectionPauseReason, type SelectionReceiver } from './selectionWatcher'

let stop: DiagramSelectionWatcher | undefined
let change: () => void
const first = createDiagramPayload('flowchart LR\nA-->B', 'png')
const second = createDiagramPayload('flowchart LR\nC-->D', 'png')

beforeEach(() => {
  vi.useFakeTimers()
  vi.spyOn(document, 'hasFocus').mockReturnValue(false)
  vi.spyOn(document, 'visibilityState', 'get').mockReturnValue('visible')
  vi.stubGlobal('Office', {
    context: {
      document: {
        addHandlerAsync: (_event: string, handler: () => void, callback: (result: { status: string }) => void) => {
          change = handler
          callback({ status: 'succeeded' })
        },
        removeHandlerAsync: vi.fn(),
      },
    },
    EventType: { DocumentSelectionChanged: 'selection-change' },
    AsyncResultStatus: { Failed: 'failed' },
  })
})

afterEach(() => {
  stop?.()
  stop = undefined
  vi.useRealTimers()
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

it('discovers and switches diagrams without any Office selection notification', async () => {
  const read = vi.fn<() => Promise<DiagramPayload | null>>().mockResolvedValue(null)
  const selected = vi.fn()
  const changed = vi.fn()
  stop = watchDiagramSelection(read, selected, vi.fn(), {
    pollIntervalMs: 500, onSelectionChange: changed,
  })
  await vi.advanceTimersByTimeAsync(0)
  read.mockResolvedValue(first)
  await vi.advanceTimersByTimeAsync(500)
  expect(selected).toHaveBeenLastCalledWith(first, 'document-poll')
  read.mockResolvedValue(second)
  await vi.advanceTimersByTimeAsync(500)
  expect(selected).toHaveBeenLastCalledWith(second, 'document-poll')
  read.mockResolvedValue(null)
  await vi.advanceTimersByTimeAsync(500)
  expect(selected).toHaveBeenLastCalledWith(null, 'document-poll')
  expect(changed).not.toHaveBeenCalled()
  await vi.advanceTimersByTimeAsync(1500)
  expect(changed).not.toHaveBeenCalled()
  expect(selected).toHaveBeenCalledTimes(4)
})

it('refreshes on pane focus for clicks shorter than the polling interval', async () => {
  const read = vi.fn().mockResolvedValue(first)
  const selected = vi.fn()
  const changed = vi.fn()
  stop = watchDiagramSelection(read, selected, vi.fn(), {
    pollIntervalMs: 500, onSelectionChange: changed,
  })
  await vi.advanceTimersByTimeAsync(0)
  vi.mocked(document.hasFocus).mockReturnValue(true)
  read.mockResolvedValue(null)
  window.dispatchEvent(new Event('focus'))
  await vi.advanceTimersByTimeAsync(0)
  expect(changed).not.toHaveBeenCalled()
  expect(selected).toHaveBeenLastCalledWith(null, 'pane-focus')
})

it('polls despite WebView focus without notifying unchanged selection', async () => {
  vi.mocked(document.hasFocus).mockReturnValue(true)
  const read = vi.fn().mockResolvedValue(null)
  const selected = vi.fn()
  const changed = vi.fn()
  stop = watchDiagramSelection(read, selected, vi.fn(), {
    pollIntervalMs: 500, onSelectionChange: changed,
  })
  await vi.advanceTimersByTimeAsync(1000)
  expect(read).toHaveBeenCalledTimes(3)
  expect(selected).toHaveBeenCalledOnce()
  expect(changed).not.toHaveBeenCalled()
  read.mockResolvedValue(first)
  await vi.advanceTimersByTimeAsync(500)
  expect(selected).toHaveBeenLastCalledWith(first, 'background-poll')
  expect(changed).not.toHaveBeenCalled()
  read.mockResolvedValue(null)
  await vi.advanceTimersByTimeAsync(500)
  expect(selected).toHaveBeenLastCalledWith(null, 'background-poll')
  // A null while typing can be caused by image replacement, not a user click.
  expect(changed).not.toHaveBeenCalled()
  read.mockResolvedValue(first)
  await vi.advanceTimersByTimeAsync(500)
  expect(selected).toHaveBeenLastCalledWith(first, 'background-poll')
  expect(changed).not.toHaveBeenCalled()
})

it('does not poll while hidden, paused, or already reading', async () => {
  let paused = false
  const read = vi.fn<() => Promise<DiagramPayload | null>>().mockResolvedValue(null)
  stop = watchDiagramSelection(read, vi.fn(), vi.fn(), {
    pollIntervalMs: 500, isPaused: () => paused,
  })
  await vi.advanceTimersByTimeAsync(0)
  vi.spyOn(document, 'visibilityState', 'get').mockReturnValue('hidden')
  await vi.advanceTimersByTimeAsync(1000)
  expect(read).toHaveBeenCalledOnce()
  vi.spyOn(document, 'visibilityState', 'get').mockReturnValue('visible')
  paused = true
  window.dispatchEvent(new Event('focus'))
  await vi.advanceTimersByTimeAsync(1000)
  expect(read).toHaveBeenCalledOnce()
  paused = false
  let finish!: (payload: DiagramPayload | null) => void
  read.mockImplementationOnce(() => new Promise(resolve => { finish = resolve }))
  await vi.advanceTimersByTimeAsync(1500)
  expect(read).toHaveBeenCalledTimes(2)
  finish(first)
  await vi.advanceTimersByTimeAsync(0)
})

it('drops stale probes and follows a newer Office selection event', async () => {
  const read = vi.fn<() => Promise<DiagramPayload | null>>().mockResolvedValue(null)
  const selected = vi.fn()
  stop = watchDiagramSelection(read, selected, vi.fn(), { pollIntervalMs: 500 })
  await vi.advanceTimersByTimeAsync(0)
  let finish!: (payload: DiagramPayload | null) => void
  read.mockImplementationOnce(() => new Promise(resolve => { finish = resolve }))
    .mockResolvedValue(second)
  await vi.advanceTimersByTimeAsync(500)
  change()
  finish(first)
  await vi.advanceTimersByTimeAsync(0)
  expect(selected).not.toHaveBeenCalledWith(first, expect.any(String))
  expect(selected).toHaveBeenLastCalledWith(second, 'office-event')
})

it('discards a probe interrupted by writing and refreshes after the write', async () => {
  let paused = false
  const read = vi.fn<() => Promise<DiagramPayload | null>>().mockResolvedValue(null)
  const selected = vi.fn()
  stop = watchDiagramSelection(read, selected, vi.fn(), {
    pollIntervalMs: 500, isPaused: () => paused,
  })
  await vi.advanceTimersByTimeAsync(0)
  let finish!: (payload: DiagramPayload | null) => void
  read.mockImplementationOnce(() => new Promise(resolve => { finish = resolve }))
    .mockResolvedValue(second)
  await vi.advanceTimersByTimeAsync(500)
  paused = true
  finish(first)
  await vi.advanceTimersByTimeAsync(1000)
  expect(selected).not.toHaveBeenCalledWith(first, expect.any(String))
  paused = false
  stop.refresh()
  await vi.advanceTimersByTimeAsync(0)
  expect(selected).toHaveBeenLastCalledWith(second, 'document-poll')
})

it('completes an Office notification queued behind an unchanged probe', async () => {
  const read = vi.fn<() => Promise<DiagramPayload | null>>().mockResolvedValue(first)
  const selected = vi.fn()
  const changed = vi.fn()
  stop = watchDiagramSelection(read, selected, vi.fn(), {
    pollIntervalMs: 500, onSelectionChange: changed,
  })
  await vi.advanceTimersByTimeAsync(0)
  let finish!: (payload: DiagramPayload | null) => void
  read.mockImplementationOnce(() => new Promise(resolve => { finish = resolve }))
  await vi.advanceTimersByTimeAsync(500)
  change()
  finish(first)
  await vi.advanceTimersByTimeAsync(0)
  expect(changed).toHaveBeenCalledExactlyOnceWith('office-event')
  expect(selected).toHaveBeenCalledTimes(2)
  expect(selected).toHaveBeenLastCalledWith(first, 'office-event')
})

it('reports lookup errors and delivers the next successful result even if unchanged', async () => {
  const read = vi.fn().mockResolvedValue(first)
  const selected = vi.fn()
  const error = vi.fn()
  stop = watchDiagramSelection(read, selected, error, { pollIntervalMs: 500 })
  await vi.advanceTimersByTimeAsync(0)
  read.mockRejectedValueOnce(new Error('Excel is unavailable'))
  await vi.advanceTimersByTimeAsync(500)
  expect(error).toHaveBeenCalledWith(expect.objectContaining({ message: 'Excel is unavailable' }))
  await vi.advanceTimersByTimeAsync(500)
  expect(selected).toHaveBeenCalledTimes(2)
  expect(selected).toHaveBeenLastCalledWith(first, 'document-poll')
})

it('cleans up polling and focus listeners and ignores in-flight results on stop', async () => {
  let finish!: (payload: DiagramPayload | null) => void
  const read = vi.fn(() => new Promise<DiagramPayload | null>(resolve => { finish = resolve }))
  const selected = vi.fn()
  stop = watchDiagramSelection(read, selected, vi.fn(), { pollIntervalMs: 500 })
  stop()
  finish(first)
  window.dispatchEvent(new Event('focus'))
  change()
  stop.refresh()
  await vi.advanceTimersByTimeAsync(2000)
  expect(read).toHaveBeenCalledOnce()
  expect(selected).not.toHaveBeenCalled()
  expect(Office.context.document.removeHandlerAsync).toHaveBeenCalledOnce()
  expect(vi.getTimerCount()).toBe(0)
})

it('leaves Word event-driven with no polling or focus refresh', async () => {
  const read = vi.fn().mockResolvedValue(null)
  stop = watchDiagramSelection(read, vi.fn(), vi.fn())
  window.dispatchEvent(new Event('focus'))
  await vi.advanceTimersByTimeAsync(2000)
  expect(read).toHaveBeenCalledOnce()
  expect(vi.getTimerCount()).toBe(0)
})

it.each(['pane-focus', 'office-event', 'document-poll'] as const)(
  'redelivers a rejected null poll when a subsequent %s confirms deselection', async (origin) => {
    vi.mocked(document.hasFocus).mockReturnValue(true)
    const read = vi.fn().mockResolvedValue(first)
    const selected = vi.fn<SelectionReceiver>((_payload, source) => source !== 'background-poll')
    stop = watchDiagramSelection(read, selected, vi.fn(), { pollIntervalMs: 500 })
    await vi.advanceTimersByTimeAsync(0)
    read.mockResolvedValue(null)
    await vi.advanceTimersByTimeAsync(500)
    expect(selected).toHaveBeenLastCalledWith(null, 'background-poll')
    if (origin === 'pane-focus') window.dispatchEvent(new Event('focus'))
    else if (origin === 'office-event') change()
    else {
      vi.mocked(document.hasFocus).mockReturnValue(false)
      await vi.advanceTimersByTimeAsync(500)
    }
    await vi.advanceTimersByTimeAsync(0)
    expect(selected).toHaveBeenLastCalledWith(null, origin)
    expect(selected).toHaveBeenCalledTimes(3)
  },
)

it('preserves the latest gesture origin across a paused read and an explicit refresh', async () => {
  let paused = false
  let finish!: (payload: DiagramPayload | null) => void
  const read = vi.fn<() => Promise<DiagramPayload | null>>().mockResolvedValue(first)
  const selected = vi.fn()
  stop = watchDiagramSelection(read, selected, vi.fn(), { pollIntervalMs: 500, isPaused: () => paused })
  await vi.advanceTimersByTimeAsync(0)
  read.mockImplementationOnce(() => new Promise(resolve => { finish = resolve })).mockResolvedValue(null)
  await vi.advanceTimersByTimeAsync(500)
  change()
  paused = true
  finish(first)
  await vi.advanceTimersByTimeAsync(0)
  expect(selected).toHaveBeenCalledOnce()
  paused = false
  stop.refresh()
  await vi.advanceTimersByTimeAsync(0)
  expect(selected).toHaveBeenLastCalledWith(null, 'office-event')
})

it('delivers an unchanged explicit refresh queued behind a document poll', async () => {
  let finish!: (payload: DiagramPayload | null) => void
  const read = vi.fn<() => Promise<DiagramPayload | null>>().mockResolvedValue(first)
  const selected = vi.fn()
  stop = watchDiagramSelection(read, selected, vi.fn(), { pollIntervalMs: 500 })
  await vi.advanceTimersByTimeAsync(0)
  read.mockImplementationOnce(() => new Promise(resolve => { finish = resolve }))
  await vi.advanceTimersByTimeAsync(500)
  stop.refresh()
  finish(first)
  await vi.advanceTimersByTimeAsync(0)
  expect(read).toHaveBeenCalledTimes(3)
  expect(selected).toHaveBeenCalledTimes(2)
  expect(selected).toHaveBeenLastCalledWith(first, 'document-poll')
})

it.each([
  ['host-write', true, 'refresh'],
  ['host-write', false, 'office-event'],
  ['settings', true, 'office-event'],
] as const)('classifies paused %s events with pane focus=%s as %s', async (reason, focused, origin) => {
  let pauseReason: SelectionPauseReason | undefined
  const read = vi.fn().mockResolvedValue(first)
  const selected = vi.fn()
  const changed = vi.fn()
  stop = watchDiagramSelection(read, selected, vi.fn(), {
    getPauseReason: () => pauseReason, onSelectionChange: changed,
  })
  await vi.advanceTimersByTimeAsync(0)
  pauseReason = reason
  vi.mocked(document.hasFocus).mockReturnValue(focused)
  read.mockResolvedValue(null)
  change()
  await vi.advanceTimersByTimeAsync(0)
  expect(read).toHaveBeenCalledOnce()
  if (origin === 'refresh') expect(changed).not.toHaveBeenCalled()
  else expect(changed).toHaveBeenCalledExactlyOnceWith('office-event')
  pauseReason = undefined
  stop.refresh()
  await vi.advanceTimersByTimeAsync(0)
  expect(selected).toHaveBeenLastCalledWith(null, origin)
})
