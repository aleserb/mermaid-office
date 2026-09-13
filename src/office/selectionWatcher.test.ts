import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { createDiagramPayload, type DiagramPayload } from '../metadata/payload'
import { watchDiagramSelection, type DiagramSelectionWatcher } from './selectionWatcher'

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
  expect(selected).toHaveBeenLastCalledWith(first)
  expect(changed).toHaveBeenCalledWith(true)
  read.mockResolvedValue(second)
  await vi.advanceTimersByTimeAsync(500)
  expect(selected).toHaveBeenLastCalledWith(second)
  read.mockResolvedValue(null)
  await vi.advanceTimersByTimeAsync(500)
  expect(selected).toHaveBeenLastCalledWith(null)
  expect(changed).toHaveBeenCalledTimes(3)
  await vi.advanceTimersByTimeAsync(1500)
  expect(changed).toHaveBeenCalledTimes(3)
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
  expect(changed).toHaveBeenCalledExactlyOnceWith(true)
  expect(selected).toHaveBeenLastCalledWith(null)
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
  expect(selected).toHaveBeenLastCalledWith(first)
  expect(changed).toHaveBeenLastCalledWith(false)
  read.mockResolvedValue(null)
  await vi.advanceTimersByTimeAsync(500)
  expect(selected).toHaveBeenLastCalledWith(null)
  // A null while typing can be caused by image replacement, not a user click.
  expect(changed).toHaveBeenLastCalledWith(false)
  read.mockResolvedValue(first)
  await vi.advanceTimersByTimeAsync(500)
  expect(selected).toHaveBeenLastCalledWith(first)
  expect(changed).toHaveBeenCalledTimes(3)
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
  expect(selected).not.toHaveBeenCalledWith(first)
  expect(selected).toHaveBeenLastCalledWith(second)
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
  expect(selected).not.toHaveBeenCalledWith(first)
  paused = false
  stop.refresh()
  await vi.advanceTimersByTimeAsync(0)
  expect(selected).toHaveBeenLastCalledWith(second)
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
  expect(changed).toHaveBeenCalledOnce()
  expect(selected).toHaveBeenCalledTimes(2)
  expect(selected).toHaveBeenLastCalledWith(first)
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
  expect(selected).toHaveBeenLastCalledWith(first)
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
