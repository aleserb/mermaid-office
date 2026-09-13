import { act, cleanup, renderHook } from '@testing-library/react'
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { renderMermaid } from '../mermaid/render'
import { DEFAULT_DIAGRAM_SETTINGS } from '../metadata/diagramSettings'
import type { DiagramDraft } from '../office/diagramRequests'
import { LIVE_UPDATE_DELAY, useDebouncedDiagramRender } from './useDebouncedDiagramRender'

vi.mock('../mermaid/render', () => ({ renderMermaid: vi.fn() }))

const draft: DiagramDraft = {
  source: 'flowchart LR\nA-->B', theme: 'default', size: 'medium', settings: DEFAULT_DIAGRAM_SETTINGS,
}

beforeEach(() => {
  vi.useFakeTimers()
  vi.mocked(renderMermaid).mockResolvedValue('<svg/>')
})

afterEach(() => {
  cleanup()
  vi.useRealTimers()
  vi.resetAllMocks()
})

it('waits for readiness and debounces source, theme and settings as one revision', async () => {
  const dispatch = vi.fn()
  const { rerender } = renderHook(({ value, ready }) => useDebouncedDiagramRender(value, ready, dispatch), {
    initialProps: { value: draft, ready: false },
  })
  await act(async () => { await vi.advanceTimersByTimeAsync(LIVE_UPDATE_DELAY) })
  expect(renderMermaid).not.toHaveBeenCalled()
  rerender({ value: draft, ready: true })
  await act(async () => { await vi.advanceTimersByTimeAsync(300) })
  const next: DiagramDraft = {
    ...draft, source: 'flowchart LR\nChanged', theme: 'dark',
    settings: { ...DEFAULT_DIAGRAM_SETTINGS, imageQuality: 'high' },
  }
  rerender({ value: next, ready: true })
  await act(async () => { await vi.advanceTimersByTimeAsync(LIVE_UPDATE_DELAY - 1) })
  expect(renderMermaid).not.toHaveBeenCalled()
  await act(async () => { await vi.advanceTimersByTimeAsync(1) })
  expect(renderMermaid).toHaveBeenCalledExactlyOnceWith(next.source, next.theme, next.settings)
  expect(dispatch).toHaveBeenCalledExactlyOnceWith({ type: 'rendered', result: { draft: next, svg: '<svg/>' } })
})

it.each(['success', 'failure'] as const)('ignores stale render %s after switching revisions', async (outcome) => {
  let finish!: (svg: string) => void
  let fail!: (error: Error) => void
  vi.mocked(renderMermaid).mockImplementationOnce(() => new Promise((resolve, reject) => {
    finish = resolve
    fail = reject
  }))
  const dispatch = vi.fn()
  const { rerender } = renderHook(({ value }) => useDebouncedDiagramRender(value, true, dispatch), {
    initialProps: { value: draft },
  })
  await act(async () => { await vi.advanceTimersByTimeAsync(LIVE_UPDATE_DELAY) })
  const next = { ...draft, source: 'flowchart LR\nCurrent' }
  rerender({ value: next })
  await act(async () => {
    if (outcome === 'success') finish('<svg>old</svg>')
    else fail(new Error('Old error'))
  })
  expect(dispatch).not.toHaveBeenCalled()
  await act(async () => { await vi.advanceTimersByTimeAsync(LIVE_UPDATE_DELAY) })
  expect(dispatch).toHaveBeenCalledExactlyOnceWith({ type: 'rendered', result: { draft: next, svg: '<svg/>' } })
})

it('normalizes current errors and cancels pending work on unmount', async () => {
  vi.mocked(renderMermaid).mockRejectedValueOnce(new Error('Parse error on line 2'))
  const dispatch = vi.fn()
  const { rerender, unmount } = renderHook(({ value }) => useDebouncedDiagramRender(value, true, dispatch), {
    initialProps: { value: draft },
  })
  await act(async () => { await vi.advanceTimersByTimeAsync(LIVE_UPDATE_DELAY) })
  expect(dispatch).toHaveBeenCalledWith({
    type: 'render-failed', draft, diagnostic: expect.objectContaining({ line: 2 }),
  })
  rerender({ value: { ...draft, source: 'flowchart LR\nLater' } })
  unmount()
  await act(async () => { await vi.advanceTimersByTimeAsync(LIVE_UPDATE_DELAY) })
  expect(renderMermaid).toHaveBeenCalledOnce()
})
