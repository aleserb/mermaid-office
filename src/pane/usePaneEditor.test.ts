import { act, cleanup, renderHook } from '@testing-library/react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createDiagramPayload } from '../metadata/payload'
import { renderMermaid } from '../mermaid/render'
import { insertDiagramWithPayload, updateDiagramById } from '../word/insertDiagram'
import { getSelectedDiagram } from '../word/selection'
import { LIVE_UPDATE_DELAY, usePaneEditor } from './usePaneEditor'

vi.mock('../mermaid/render', () => ({ renderMermaid: vi.fn() }))
vi.mock('../word/insertDiagram', () => ({
  insertDiagramWithPayload: vi.fn(),
  updateDiagramById: vi.fn(),
}))
vi.mock('../word/selection', () => ({ getSelectedDiagram: vi.fn() }))

const existing = createDiagramPayload('flowchart LR\nA-->B', 'png', 'forest')

beforeEach(() => {
  vi.useFakeTimers()
  vi.stubGlobal('Office', { onReady: vi.fn().mockResolvedValue({}) })
  vi.mocked(getSelectedDiagram).mockResolvedValue(null)
  vi.mocked(renderMermaid).mockImplementation(async (source) => `<svg>${source}</svg>`)
  vi.mocked(updateDiagramById).mockResolvedValue('png')
  vi.mocked(insertDiagramWithPayload).mockImplementation(async (_svg, source, theme, size) =>
    createDiagramPayload(source, 'png', theme, size))
})

afterEach(() => {
  cleanup()
  vi.useRealTimers()
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
  vi.resetAllMocks()
  localStorage.clear()
})

async function openPane(payload: typeof existing | null = null) {
  vi.mocked(getSelectedDiagram).mockResolvedValue(payload)
  const hook = renderHook(usePaneEditor)
  await act(async () => {})
  return hook
}

async function renderPending() {
  await act(async () => { await vi.advanceTimersByTimeAsync(LIVE_UPDATE_DELAY) })
}

describe('code-only pane workflow', () => {
  it('requires explicit insertion and keeps the returned diagram as its update target', async () => {
    const { result } = await openPane()
    await renderPending()
    expect(result.current.canInsert).toBe(true)
    expect(insertDiagramWithPayload).not.toHaveBeenCalled()
    expect(updateDiagramById).not.toHaveBeenCalled()

    await act(async () => { await result.current.insert() })
    const insertedId = result.current.target?.id
    expect(insertedId).toBeTruthy()
    expect(result.current.canInsert).toBe(false)
    act(() => result.current.changeSource('flowchart LR\nA-->C'))
    await renderPending()
    expect(updateDiagramById).toHaveBeenCalledWith(
      expect.any(String), expect.objectContaining({ id: insertedId }),
      'flowchart LR\nA-->C', 'redux-color', 'medium', false,
    )
  })

  it('debounces typing and updates the pinned diagram, not a newly selected one', async () => {
    const { result } = await openPane(existing)
    await renderPending()
    expect(updateDiagramById).not.toHaveBeenCalled()
    vi.mocked(getSelectedDiagram).mockResolvedValue(createDiagramPayload('flowchart LR\nOther', 'png'))
    act(() => result.current.changeSource('flowchart LR\nA-->C'))
    await act(async () => { await vi.advanceTimersByTimeAsync(300) })
    act(() => result.current.changeSource('flowchart LR\nA-->D'))
    await renderPending()
    expect(updateDiagramById).toHaveBeenCalledOnce()
    expect(updateDiagramById).toHaveBeenCalledWith(
      expect.any(String), existing, 'flowchart LR\nA-->D', 'forest', 'medium', false,
    )
    expect(getSelectedDiagram).toHaveBeenCalledOnce()
  })

  it('keeps the last valid image when the source has errors', async () => {
    const { result } = await openPane(existing)
    await renderPending()
    vi.mocked(renderMermaid).mockRejectedValue(new Error('Parse error on line 2'))
    act(() => result.current.changeSource('broken'))
    await renderPending()
    expect(result.current.diagnostic?.message).toContain('Line 2')
    expect(updateDiagramById).not.toHaveBeenCalled()
    expect(result.current.target?.source).toBe(existing.source)
  })

  it('ignores an old render that finishes after newer source', async () => {
    const { result } = await openPane(existing)
    await renderPending()
    let finishOlder: (value: string) => void = () => { throw new Error('Render has not started') }
    vi.mocked(renderMermaid).mockImplementationOnce(() => new Promise((resolve) => { finishOlder = resolve }))
    act(() => result.current.changeSource('flowchart LR\nOld'))
    await renderPending()
    act(() => result.current.changeSource('flowchart LR\nCurrent'))
    await renderPending()
    await act(async () => { finishOlder('<svg>old</svg>') })
    expect(updateDiagramById).toHaveBeenCalledOnce()
    expect(result.current.target?.source).toBe('flowchart LR\nCurrent')
  })

  it('applies theme changes to the linked Word diagram', async () => {
    const { result } = await openPane(existing)
    await renderPending()
    act(() => result.current.changeTheme('redux-color'))
    await renderPending()
    expect(updateDiagramById).toHaveBeenCalledWith(
      expect.any(String), existing, existing.source, 'redux-color', 'medium', false,
    )
    expect(localStorage.getItem('mermaid-office:preferred-theme')).toBe('redux-color')
  })

  it('tracks the new format when live-updating a legacy SVG diagram', async () => {
    const { result } = await openPane({ ...existing, format: 'svg' })
    act(() => result.current.changeSource('flowchart LR\nUpdated'))
    await renderPending()
    expect(result.current.target?.format).toBe('png')
    act(() => result.current.changeSource('flowchart LR\nUpdatedAgain'))
    await renderPending()
    expect(updateDiagramById).toHaveBeenLastCalledWith(
      expect.any(String), expect.objectContaining({ format: 'png' }),
      'flowchart LR\nUpdatedAgain', 'forest', 'medium', false,
    )
  })

  it('serializes writes and does not discard typing while a write is in flight', async () => {
    const { result } = await openPane(existing)
    await renderPending()
    let finish: (value: 'png') => void = () => { throw new Error('Write has not started') }
    vi.mocked(updateDiagramById).mockImplementationOnce(() => new Promise((resolve) => { finish = resolve }))
    act(() => result.current.changeSource('flowchart LR\nFirst'))
    await renderPending()
    expect(result.current.writing).toBe(true)
    act(() => result.current.changeSource('flowchart LR\nLatest'))
    await renderPending()
    expect(updateDiagramById).toHaveBeenCalledOnce()

    await act(async () => { finish('png') })
    expect(updateDiagramById).toHaveBeenCalledTimes(2)
    expect(result.current.draft.source).toBe('flowchart LR\nLatest')
    expect(result.current.target?.source).toBe('flowchart LR\nLatest')
    expect(result.current.dirty).toBe(false)
  })

  it('retains edits typed during initial insertion and then updates the inserted diagram', async () => {
    const { result } = await openPane()
    await renderPending()
    const inserted = createDiagramPayload(result.current.draft.source, 'png', 'redux-color')
    let finish: (value: typeof inserted) => void = () => { throw new Error('Insert has not started') }
    vi.mocked(insertDiagramWithPayload).mockImplementationOnce(() => new Promise((resolve) => { finish = resolve }))
    act(() => { void result.current.insert() })
    await act(async () => {})
    act(() => result.current.changeSource('flowchart LR\nTypedDuringInsert'))
    await renderPending()
    await act(async () => { finish(inserted) })
    expect(result.current.draft.source).toContain('TypedDuringInsert')
    expect(updateDiagramById).toHaveBeenCalledWith(
      expect.any(String), inserted, 'flowchart LR\nTypedDuringInsert', 'redux-color', 'medium', false,
    )
  })

  it('reports failed writes without retrying in a loop', async () => {
    vi.spyOn(console, 'error').mockImplementation(() => {})
    const { result } = await openPane(existing)
    await renderPending()
    vi.mocked(updateDiagramById).mockRejectedValueOnce(new Error('The diagram was deleted.'))
    act(() => result.current.changeSource('flowchart LR\nChanged'))
    await renderPending()
    await renderPending()
    expect(updateDiagramById).toHaveBeenCalledOnce()
    expect(result.current.wordError).toContain('deleted')
    expect(result.current.canRetry).toBe(true)
    act(() => result.current.retry())
    await act(async () => {})
    expect(updateDiagramById).toHaveBeenCalledTimes(2)
  })

  it('confirms switching away from unsaved edits and pauses pending live writes', async () => {
    const { result } = await openPane(existing)
    await renderPending()
    act(() => result.current.changeSource('flowchart LR\nUnsaved'))
    act(() => result.current.newDiagram())
    expect(result.current.pending).not.toBeNull()
    await renderPending()
    expect(updateDiagramById).not.toHaveBeenCalled()
    act(() => result.current.discardAndSwitch())
    expect(result.current.target).toBeNull()
    expect(result.current.draft.source).not.toContain('Unsaved')
  })

  it('loads another diagram only on an explicit Edit selected action', async () => {
    const { result } = await openPane(existing)
    const other = createDiagramPayload('flowchart LR\nOther', 'png', 'neutral')
    vi.mocked(getSelectedDiagram).mockResolvedValue(other)
    await act(async () => { await result.current.loadSelected() })
    expect(result.current.target).toEqual(other)
    expect(result.current.draft.source).toBe(other.source)
    expect(updateDiagramById).not.toHaveBeenCalled()
  })

  it('does not lose typing while a selected-diagram lookup is in flight', async () => {
    const { result } = await openPane(existing)
    const other = createDiagramPayload('flowchart LR\nOther', 'png')
    let finish: (value: typeof other) => void = () => { throw new Error('Lookup has not started') }
    vi.mocked(getSelectedDiagram).mockImplementationOnce(() => new Promise((resolve) => { finish = resolve }))
    act(() => { void result.current.loadSelected() })
    act(() => result.current.changeSource('flowchart LR\nKeepTheseEdits'))
    await act(async () => { finish(other) })
    expect(result.current.pending?.target).toEqual(other)
    expect(result.current.draft.source).toContain('KeepTheseEdits')
    act(() => result.current.keepEditing())
    expect(result.current.target?.id).toBe(existing.id)
  })

  it('does not overwrite a selected diagram when inserting a new draft', async () => {
    const { result } = await openPane()
    await renderPending()
    vi.mocked(getSelectedDiagram).mockResolvedValue(existing)
    await act(async () => { await result.current.insert() })
    expect(insertDiagramWithPayload).not.toHaveBeenCalled()
    expect(result.current.wordError).toContain('blank line')
  })

  it('cancels pending rendering when the pane is closed', async () => {
    const { result, unmount } = await openPane(existing)
    act(() => result.current.changeSource('flowchart LR\nNotSaved'))
    unmount()
    await renderPending()
    expect(updateDiagramById).not.toHaveBeenCalled()
  })
})
