import { act, cleanup, renderHook } from '@testing-library/react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createDiagramPayload } from '../metadata/payload'
import { DEFAULT_DIAGRAM_SETTINGS } from '../metadata/diagramSettings'
import { DEFAULT_DIAGRAM } from '../defaultDiagram'
import { renderMermaid } from '../mermaid/render'
import { insertDiagramWithPayload, updateDiagramById } from '../word/insertDiagram'
import { getSelectedDiagram, watchSelectedDiagram } from '../word/selection'
import { LIVE_UPDATE_DELAY, usePaneEditor } from './usePaneEditor'
import { watchDiagramSelection, type SelectionOrigin, type SelectionReceiver } from '../office/selectionWatcher'

vi.mock('../mermaid/render', async importOriginal => ({
  ...await importOriginal<typeof import('../mermaid/render')>(),
  renderMermaid: vi.fn(),
}))
vi.mock('../word/insertDiagram', () => ({
  insertDiagramWithPayload: vi.fn(),
  updateDiagramById: vi.fn(),
}))
vi.mock('../word/selection', () => ({ getSelectedDiagram: vi.fn(), watchSelectedDiagram: vi.fn() }))

const existing = createDiagramPayload('flowchart LR\nA-->B', 'png', 'forest')
const large = createDiagramPayload(`flowchart LR\n${'A-->B\n'.repeat(50)}`, 'png', 'forest')
let selected: typeof existing | null = null
let receiveSelection: SelectionReceiver
let notifySelection: ((origin: SelectionOrigin) => void) | undefined
let isPaused: (() => boolean) | undefined
let queuedOrigin: SelectionOrigin
const stopWatching = vi.fn()
const refreshSelection = vi.fn()

function selectInWord(payload: typeof existing | null, origin: SelectionOrigin = 'office-event') {
  selected = payload
  queuedOrigin = origin
  if (origin === 'office-event') notifySelection?.(origin)
  if (!isPaused?.()) {
    receiveSelection(payload, origin)
    queuedOrigin = 'refresh'
  }
}

beforeEach(() => {
  vi.useFakeTimers()
  vi.stubGlobal('Office', { onReady: vi.fn().mockResolvedValue({}), context: { document: {} } })
  vi.spyOn(document, 'hasFocus').mockReturnValue(true)
  vi.mocked(getSelectedDiagram).mockResolvedValue(null)
  queuedOrigin = 'refresh'
  vi.mocked(watchSelectedDiagram).mockImplementation((onSelected, _onError, options) => {
    receiveSelection = onSelected
    notifySelection = options?.onSelectionChange
    isPaused = options?.isPaused
    refreshSelection.mockImplementation(() => {
      if (!isPaused?.()) {
        onSelected(selected, queuedOrigin)
        queuedOrigin = 'refresh'
      }
    })
    void getSelectedDiagram().then(payload => onSelected(payload, 'initial'))
    return Object.assign(stopWatching, { refresh: refreshSelection })
  })
  vi.mocked(renderMermaid).mockImplementation(async (source) => `<svg>${source}</svg>`)
  vi.mocked(updateDiagramById).mockResolvedValue('png')
  vi.mocked(insertDiagramWithPayload).mockImplementation(async ({ draft }) =>
    createDiagramPayload(draft.source, 'png', draft.theme, draft.size, draft.settings))
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
  selected = payload
  vi.mocked(getSelectedDiagram).mockResolvedValue(payload)
  const hook = renderHook(usePaneEditor)
  await act(async () => {})
  return hook
}

async function renderPending() {
  await act(async () => { await vi.advanceTimersByTimeAsync(LIVE_UPDATE_DELAY) })
}

function useRealSelectionWatcher() {
  let officeSelectionChanged!: () => void
  vi.stubGlobal('Office', {
    onReady: vi.fn().mockResolvedValue({}),
    context: {
      document: {
        addHandlerAsync: (_event: string, handler: () => void, callback: (result: { status: string }) => void) => {
          officeSelectionChanged = handler
          callback({ status: 'succeeded' })
        },
        removeHandlerAsync: vi.fn(),
      },
    },
    EventType: { DocumentSelectionChanged: 'selection-change' },
    AsyncResultStatus: { Failed: 'failed' },
  })
  // Use the real watcher with the Word adapter's mocked writer so these tests
  // exercise the editor/watcher boundary independently of either host's reader.
  vi.mocked(watchSelectedDiagram).mockImplementation((onSelected, onError, options) =>
    watchDiagramSelection(() => Promise.resolve(selected), onSelected, onError, { ...options, pollIntervalMs: 500 }))
  vi.spyOn(document, 'visibilityState', 'get').mockReturnValue('visible')
  return () => officeSelectionChanged()
}

describe('code-only pane workflow', () => {
  it.each(['pane-focus', 'office-event'] as const)(
    'does not deduplicate a real %s deselection against a rejected focused null poll', async (origin) => {
      const officeSelectionChanged = useRealSelectionWatcher()
      const { result } = await openPane(existing)
      const history = result.current.historyKey
      selected = null
      await act(async () => { await vi.advanceTimersByTimeAsync(500) })
      expect(result.current.target?.id).toBe(existing.id)
      expect(result.current.historyKey).toBe(history)
      await act(async () => {
        if (origin === 'pane-focus') window.dispatchEvent(new Event('focus'))
        else officeSelectionChanged()
      })
      expect(document.hasFocus()).toBe(true)
      expect(result.current.target).toBeNull()
      expect(result.current.historyKey).toBe(history + 1)
      expect(result.current.loadingSelection).toBe(false)
    },
  )

  it.each(['pane-focus', 'office-event'] as const)(
    'confirms dirty deselection on %s after rejecting a focused null poll', async (origin) => {
      const officeSelectionChanged = useRealSelectionWatcher()
      const { result } = await openPane(existing)
      const history = result.current.historyKey
      act(() => result.current.changeSource('flowchart LR\nUnsaved'))
      selected = null
      await act(async () => { await vi.advanceTimersByTimeAsync(500) })
      expect(result.current.pending).toBeNull()
      await act(async () => {
        if (origin === 'pane-focus') window.dispatchEvent(new Event('focus'))
        else officeSelectionChanged()
      })
      expect(result.current.pending).toEqual({ target: null })
      expect(result.current.draft.source).toContain('Unsaved')
      expect(result.current.historyKey).toBe(history)
      await renderPending()
      expect(updateDiagramById).not.toHaveBeenCalled()
    },
  )

  it('preserves undo history and dismissed confirmation across unchanged polls and refreshes', async () => {
    const officeSelectionChanged = useRealSelectionWatcher()
    const { result } = await openPane(large)
    const history = result.current.historyKey
    act(() => result.current.changeSource(`${large.source}B-->C`))
    await act(async () => { await vi.advanceTimersByTimeAsync(1500) })
    expect(result.current.historyKey).toBe(history)
    expect(result.current.pending).toBeNull()
    selected = null
    await act(async () => officeSelectionChanged())
    act(() => result.current.keepEditing())
    await act(async () => { await vi.advanceTimersByTimeAsync(1500) })
    expect(result.current.pending).toBeNull()
    expect(result.current.historyKey).toBe(history)
    expect(result.current.draft.source).toBe(`${large.source}B-->C`)
  })

  it('honors a document poll after insertion even when the last accepted snapshot was empty', async () => {
    useRealSelectionWatcher()
    const { result } = await openPane()
    await renderPending()
    await act(async () => result.current.insert())
    expect(result.current.target).not.toBeNull()
    vi.mocked(document.hasFocus).mockReturnValue(false)
    await act(async () => { await vi.advanceTimersByTimeAsync(500) })
    expect(result.current.target).toBeNull()
  })

  it.each(['PC', 'Mac'] as const)('keeps live updates enabled for large diagrams on desktop %s', async (platform) => {
    vi.stubGlobal('Office', {
      onReady: vi.fn().mockResolvedValue({}),
      context: { document: {}, platform },
    })
    const { result } = await openPane(large)
    await renderPending()
    expect(result.current.manualUpdates).toBe(false)
    act(() => result.current.changeSource(`${large.source}B-->C`))
    await renderPending()
    expect(updateDiagramById).toHaveBeenCalledOnce()
    expect(result.current.manualUpdates).toBe(false)
  })

  it('keeps large diagrams unchanged until Update and writes only the requested revision', async () => {
    const { result } = await openPane(large)
    await renderPending()
    expect(result.current.manualUpdates).toBe(true)
    expect(result.current.canUpdate).toBe(false)
    act(() => result.current.changeSource(`${large.source}B-->C`))
    expect(result.current.canUpdate).toBe(false)
    await renderPending()
    expect(result.current.canUpdate).toBe(true)
    expect(updateDiagramById).not.toHaveBeenCalled()

    let finish!: (value: 'png') => void
    vi.mocked(updateDiagramById).mockImplementationOnce(() => new Promise(resolve => { finish = resolve }))
    act(() => result.current.update())
    expect(result.current.writing).toBe(true)
    act(() => result.current.changeSource(`${large.source}B-->D`))
    await renderPending()
    await act(async () => { finish('png') })
    expect(updateDiagramById).toHaveBeenCalledOnce()
    expect(result.current.target?.source).toBe(`${large.source}B-->C`)
    expect(result.current.dirty).toBe(true)
    expect(result.current.canUpdate).toBe(true)
    await act(async () => result.current.update())
    expect(updateDiagramById).toHaveBeenCalledTimes(2)
    expect(result.current.target?.source).toBe(`${large.source}B-->D`)
    expect(result.current.canUpdate).toBe(false)
  })

  it('switches to manual updates before writing a growing diagram and stays manual when shortened', async () => {
    const { result } = await openPane(existing)
    act(() => result.current.changeSource(large.source))
    await renderPending()
    expect(result.current.manualUpdates).toBe(true)
    expect(updateDiagramById).not.toHaveBeenCalled()
    act(() => result.current.changeSource('flowchart LR\nSmall'))
    await renderPending()
    expect(result.current.manualUpdates).toBe(true)
    expect(updateDiagramById).not.toHaveBeenCalled()
    await act(async () => result.current.update())
    act(() => selectInWord({ ...existing, id: 'another-small-diagram' }))
    expect(result.current.manualUpdates).toBe(false)
  })

  it('uses rendered dimensions to detect a large diagram with short source before writing', async () => {
    const { result } = await openPane(existing)
    await renderPending()
    vi.mocked(renderMermaid).mockResolvedValueOnce('<svg viewBox="0 0 300 2000"/>')
    act(() => result.current.changeSource('flowchart TD\nTall'))
    await renderPending()
    expect(result.current.manualUpdates).toBe(true)
    expect(result.current.canUpdate).toBe(true)
    expect(updateDiagramById).not.toHaveBeenCalled()
  })

  it('retains explicit insertion for large drafts and requires Update for subsequent edits', async () => {
    const { result } = await openPane()
    act(() => result.current.changeSource(large.source))
    await renderPending()
    await act(async () => { await result.current.insert() })
    expect(insertDiagramWithPayload).toHaveBeenCalledOnce()
    act(() => result.current.changeSource(`${large.source}B-->C`))
    await renderPending()
    expect(result.current.canUpdate).toBe(true)
    expect(updateDiagramById).not.toHaveBeenCalled()
  })

  it('blocks invalid manual updates and requires another click after a failed write', async () => {
    vi.spyOn(console, 'error').mockImplementation(() => {})
    const { result } = await openPane(large)
    vi.mocked(renderMermaid).mockRejectedValueOnce(new Error('Invalid Mermaid'))
    act(() => result.current.changeSource('invalid'))
    await renderPending()
    expect(result.current.canUpdate).toBe(false)
    act(() => result.current.update())
    expect(updateDiagramById).not.toHaveBeenCalled()
    act(() => result.current.changeSource(`${large.source}B-->C`))
    await renderPending()
    vi.mocked(updateDiagramById).mockRejectedValueOnce(new Error('Word is unavailable'))
    await act(async () => result.current.update())
    await renderPending()
    expect(result.current.canRetry).toBe(true)
    expect(updateDiagramById).toHaveBeenCalledOnce()
    await act(async () => result.current.update())
    expect(updateDiagramById).toHaveBeenCalledTimes(2)
    expect(result.current.canRetry).toBe(false)
  })

  it('confirms switching away from manual edits and disables Update during confirmation', async () => {
    const { result } = await openPane(large)
    act(() => result.current.changeSource(`${large.source}B-->C`))
    await renderPending()
    act(() => selectInWord(null))
    expect(result.current.pending).not.toBeNull()
    expect(result.current.canUpdate).toBe(false)
    expect(updateDiagramById).not.toHaveBeenCalled()
    act(() => result.current.keepEditing())
    expect(result.current.canUpdate).toBe(true)
  })

  it('treats settings Apply as an explicit update for the pinned large diagram', async () => {
    const { result } = await openPane(large)
    await renderPending()
    act(() => result.current.beginSettings())
    const other = { ...existing, id: 'other-diagram' }
    act(() => selectInWord(other))
    act(() => result.current.applySettings('dark', DEFAULT_DIAGRAM_SETTINGS))
    await renderPending()
    expect(updateDiagramById).not.toHaveBeenCalled()
    await act(async () => result.current.endSettings(true))
    expect(updateDiagramById).toHaveBeenCalledExactlyOnceWith(
      { svg: expect.any(String), existing: large,
        draft: { source: large.source, theme: 'dark', size: 'medium', settings: DEFAULT_DIAGRAM_SETTINGS },
        applySize: false },
    )
    expect(result.current.target?.id).toBe(other.id)
    expect(isPaused?.()).toBe(false)
  })

  it('unpins selection if typing cancels the manual revision requested by settings Apply', async () => {
    const { result } = await openPane(large)
    await renderPending()
    act(() => result.current.beginSettings())
    act(() => result.current.applySettings('dark', DEFAULT_DIAGRAM_SETTINGS))
    act(() => result.current.endSettings(true))
    expect(isPaused?.()).toBe(true)
    act(() => result.current.changeSource(`${large.source}B-->C`))
    expect(isPaused?.()).toBe(false)
    await renderPending()
    expect(result.current.canUpdate).toBe(true)
    expect(updateDiagramById).not.toHaveBeenCalled()
  })

  it('pins settings to the original diagram and applies before following a new selection', async () => {
    const { result } = await openPane(existing)
    await renderPending()
    act(() => result.current.beginSettings())
    const other = { ...existing, id: 'other-diagram', theme: 'neutral' as const }
    act(() => selectInWord(other))
    expect(isPaused?.()).toBe(true)
    expect(result.current.target?.id).toBe(existing.id)
    expect(result.current.loadingSelection).toBe(false)
    const settings = { ...DEFAULT_DIAGRAM_SETTINGS, imageQuality: 'high' as const }
    act(() => result.current.applySettings('dark', settings))
    await renderPending()
    expect(updateDiagramById).not.toHaveBeenCalled()
    await act(async () => { result.current.endSettings(true) })
    await renderPending()
    expect(updateDiagramById).toHaveBeenCalledExactlyOnceWith(
      { svg: expect.any(String), existing,
        draft: { source: existing.source, theme: 'dark', size: 'medium', settings }, applySize: false },
    )
    expect(result.current.target?.id).toBe(other.id)
    expect(result.current.settingsActive).toBe(false)
    expect(isPaused?.()).toBe(false)
  })

  it('resumes selection on settings Cancel without applying to either diagram', async () => {
    const { result } = await openPane(existing)
    act(() => result.current.beginSettings())
    const other = { ...existing, id: 'other-diagram' }
    act(() => selectInWord(other))
    act(() => result.current.endSettings(false))
    expect(result.current.target?.id).toBe(other.id)
    await renderPending()
    expect(updateDiagramById).not.toHaveBeenCalled()
  })

  it('does not remain pinned when applying unchanged settings after a failed write', async () => {
    vi.spyOn(console, 'error').mockImplementation(() => {})
    const { result } = await openPane(existing)
    vi.mocked(updateDiagramById).mockRejectedValueOnce(new Error('Word write failed'))
    act(() => result.current.changeSource('flowchart LR\nChanged'))
    await renderPending()
    expect(result.current.canRetry).toBe(true)
    act(() => result.current.beginSettings())
    act(() => result.current.applySettings(existing.theme, DEFAULT_DIAGRAM_SETTINGS))
    await act(async () => { result.current.endSettings(true) })
    expect(isPaused?.()).toBe(false)
    expect(result.current.canRetry).toBe(true)
  })

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
      { svg: expect.any(String), existing: expect.objectContaining({ id: insertedId }),
        draft: { source: 'flowchart LR\nA-->C', theme: 'redux-color', size: 'medium', settings: DEFAULT_DIAGRAM_SETTINGS },
        applySize: false },
    )
  })

  it('debounces typing without reloading the same diagram after a live update', async () => {
    const { result } = await openPane(existing)
    await renderPending()
    expect(updateDiagramById).not.toHaveBeenCalled()
    act(() => result.current.changeSource('flowchart LR\nA-->C'))
    await act(async () => { await vi.advanceTimersByTimeAsync(300) })
    act(() => result.current.changeSource('flowchart LR\nA-->D'))
    await renderPending()
    expect(updateDiagramById).toHaveBeenCalledOnce()
    expect(updateDiagramById).toHaveBeenCalledWith(
      { svg: expect.any(String), existing,
        draft: { source: 'flowchart LR\nA-->D', theme: 'forest', size: 'medium', settings: DEFAULT_DIAGRAM_SETTINGS },
        applySize: false },
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
    await act(async () => { finishOlder('<svg viewBox="0 0 300 2000"/>') })
    expect(result.current.manualUpdates).toBe(false)
    expect(updateDiagramById).toHaveBeenCalledOnce()
    expect(result.current.target?.source).toBe('flowchart LR\nCurrent')
  })

  it('applies theme changes to the linked Word diagram', async () => {
    const { result } = await openPane(existing)
    await renderPending()
    act(() => result.current.changeTheme('redux-color'))
    await renderPending()
    expect(updateDiagramById).toHaveBeenCalledWith(
      { svg: expect.any(String), existing,
        draft: { source: existing.source, theme: 'redux-color', size: 'medium', settings: DEFAULT_DIAGRAM_SETTINGS },
        applySize: false },
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
      { svg: expect.any(String), existing: expect.objectContaining({ format: 'png' }),
        draft: { source: 'flowchart LR\nUpdatedAgain', theme: 'forest', size: 'medium', settings: DEFAULT_DIAGRAM_SETTINGS },
        applySize: false },
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
      { svg: expect.any(String), existing: inserted,
        draft: { source: 'flowchart LR\nTypedDuringInsert', theme: 'redux-color', size: 'medium', settings: DEFAULT_DIAGRAM_SETTINGS },
        applySize: false },
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
    expect(result.current.hostError).toContain('deleted')
    expect(result.current.canRetry).toBe(true)
    act(() => result.current.retry())
    await act(async () => {})
    expect(updateDiagramById).toHaveBeenCalledTimes(2)
  })

  it('confirms switching away from unsaved edits and pauses pending live writes', async () => {
    const { result } = await openPane(existing)
    await renderPending()
    act(() => result.current.changeSource('flowchart LR\nUnsaved'))
    act(() => selectInWord(null))
    expect(result.current.pending).not.toBeNull()
    await renderPending()
    expect(updateDiagramById).not.toHaveBeenCalled()
    act(() => result.current.discardAndSwitch())
    expect(result.current.target).toBeNull()
    expect(result.current.draft.source).not.toContain('Unsaved')
  })

  it('automatically loads a newly selected diagram without a button press', async () => {
    const { result } = await openPane(existing)
    const other = createDiagramPayload('flowchart LR\nOther', 'png', 'neutral')
    act(() => selectInWord(other))
    expect(result.current.target).toEqual(other)
    expect(result.current.draft.source).toBe(other.source)
    expect(updateDiagramById).not.toHaveBeenCalled()
  })

  it('uses default code and enables Insert when no diagram is selected', async () => {
    const { result } = await openPane(existing)
    act(() => selectInWord(null))
    expect(result.current.target).toBeNull()
    expect(result.current.draft.source).toBe(DEFAULT_DIAGRAM)
    expect(result.current.draft.theme).toBe('redux-color')
    await renderPending()
    expect(result.current.canInsert).toBe(true)
  })

  it('does not reset a new draft when moving between blank lines', async () => {
    const { result } = await openPane()
    act(() => result.current.changeSource('flowchart LR\nNewDraft'))
    act(() => selectInWord(null))
    expect(result.current.draft.source).toContain('NewDraft')
    expect(result.current.pending).toBeNull()
  })

  it('can reselect the same picture after moving to a blank line', async () => {
    const { result } = await openPane(existing)
    act(() => selectInWord(null))
    expect(result.current.target).toBeNull()
    act(() => selectInWord(existing))
    expect(result.current.target?.id).toBe(existing.id)
    expect(result.current.draft.source).toBe(existing.source)
  })

  it('preserves typing and undo history for repeated selection notifications', async () => {
    const { result } = await openPane(existing)
    const history = result.current.historyKey
    act(() => result.current.changeSource('flowchart LR\nKeepTheseEdits'))
    act(() => selectInWord(existing, 'refresh'))
    expect(result.current.draft.source).toContain('KeepTheseEdits')
    expect(result.current.historyKey).toBe(history)
    expect(result.current.pending).toBeNull()
  })

  it('does not reset the pane when its own replacement clears document selection', async () => {
    const { result } = await openPane(existing)
    act(() => selectInWord(null, 'refresh'))
    expect(result.current.target?.id).toBe(existing.id)
    expect(result.current.draft.source).toBe(existing.source)
  })

  it('honors deselection detected after focus returns to the pane', async () => {
    const { result } = await openPane(existing)
    expect(document.hasFocus()).toBe(true)
    act(() => {
      receiveSelection(null, 'pane-focus')
    })
    expect(result.current.target).toBeNull()
    expect(result.current.loadingSelection).toBe(false)
  })

  it('preserves unsaved edits when a focus refresh detects deselection', async () => {
    const { result } = await openPane(existing)
    act(() => result.current.changeSource('flowchart LR\nUnsaved'))
    act(() => {
      receiveSelection(null, 'pane-focus')
    })
    expect(result.current.target?.id).toBe(existing.id)
    expect(result.current.draft.source).toBe('flowchart LR\nUnsaved')
    expect(result.current.pending).toEqual({ target: null })
  })

  it('preserves unsaved edits and confirms a selection-driven switch', async () => {
    const { result } = await openPane(existing)
    const other = createDiagramPayload('flowchart LR\nOther', 'png')
    act(() => result.current.changeSource('flowchart LR\nKeepTheseEdits'))
    act(() => selectInWord(other))
    expect(result.current.pending?.target).toEqual(other)
    expect(result.current.draft.source).toContain('KeepTheseEdits')
    act(() => result.current.keepEditing())
    expect(result.current.target?.id).toBe(existing.id)
  })

  it('uses the latest selection when confirming a pending switch', async () => {
    const { result } = await openPane(existing)
    act(() => result.current.changeSource('flowchart LR\nUnsaved'))
    act(() => selectInWord(createDiagramPayload('flowchart LR\nOther', 'png')))
    act(() => selectInWord(null))
    expect(result.current.pending).toEqual({ target: null })
    await renderPending()
    expect(updateDiagramById).not.toHaveBeenCalled()
    act(() => result.current.discardAndSwitch())
    expect(result.current.target).toBeNull()
    expect(result.current.draft.source).toBe(DEFAULT_DIAGRAM)
  })

  it('cancels a pending switch when the user reselects the current diagram', async () => {
    const { result } = await openPane(existing)
    act(() => result.current.changeSource('flowchart LR\nKeepTheseEdits'))
    act(() => selectInWord(null))
    act(() => selectInWord(existing))
    expect(result.current.pending).toBeNull()
    expect(result.current.draft.source).toContain('KeepTheseEdits')
  })

  it('switches to the latest document selection after an in-flight write finishes', async () => {
    const { result } = await openPane(existing)
    let finish: (value: 'png') => void = () => { throw new Error('Write has not started') }
    vi.mocked(updateDiagramById).mockImplementationOnce(() => new Promise((resolve) => { finish = resolve }))
    act(() => result.current.changeSource('flowchart LR\nSavedBeforeSwitch'))
    await renderPending()
    const other = createDiagramPayload('flowchart LR\nOther', 'png')
    act(() => selectInWord(other))
    expect(result.current.target?.id).toBe(existing.id)
    await act(async () => { finish('png') })
    expect(result.current.target?.id).toBe(other.id)
    expect(result.current.draft.source).toBe(other.source)
    expect(updateDiagramById).toHaveBeenCalledOnce()
  })

  it('returns to Insert after a real deselection during an in-flight write', async () => {
    const { result } = await openPane(existing)
    let finish: (value: 'png') => void = () => { throw new Error('Write has not started') }
    vi.mocked(updateDiagramById).mockImplementationOnce(() => new Promise((resolve) => { finish = resolve }))
    act(() => result.current.changeSource('flowchart LR\nSaved'))
    await renderPending()
    act(() => selectInWord(null))
    await act(async () => { finish('png') })
    expect(result.current.target).toBeNull()
    expect(result.current.draft.source).toBe(DEFAULT_DIAGRAM)
    await renderPending()
    expect(result.current.canInsert).toBe(true)
  })

  it('does not overwrite a selected diagram when inserting a new draft', async () => {
    const { result } = await openPane()
    await renderPending()
    vi.mocked(getSelectedDiagram).mockResolvedValue(existing)
    await act(async () => { await result.current.insert() })
    expect(insertDiagramWithPayload).not.toHaveBeenCalled()
    expect(result.current.hostError).toContain('blank line')
  })

  it('cancels pending rendering when the pane is closed', async () => {
    const { result, unmount } = await openPane(existing)
    act(() => result.current.changeSource('flowchart LR\nNotSaved'))
    unmount()
    await renderPending()
    expect(updateDiagramById).not.toHaveBeenCalled()
    expect(stopWatching).toHaveBeenCalledOnce()
  })

  it('applies settings and theme atomically and saves even a settings-only edit', async () => {
    const { result } = await openPane(existing)
    await renderPending()
    const settings = { ...DEFAULT_DIAGRAM_SETTINGS, fontSize: 20 as const, imageQuality: 'high' as const }
    act(() => result.current.applySettings(existing.theme, settings))
    expect(result.current.dirty).toBe(true)
    expect(updateDiagramById).not.toHaveBeenCalled()
    await renderPending()
    expect(renderMermaid).toHaveBeenLastCalledWith(existing.source, existing.theme, settings)
    expect(updateDiagramById).toHaveBeenCalledExactlyOnceWith(
      { svg: expect.any(String), existing,
        draft: { source: existing.source, theme: existing.theme, size: existing.size, settings },
        applySize: false },
    )
    expect(result.current.target?.settings).toEqual(settings)
    expect(result.current.dirty).toBe(false)
    expect(JSON.parse(localStorage.getItem('mermaid-office:preferred-settings')!)).toEqual(settings)
  })

  it('restores selected settings and keeps legacy diagrams on defaults', async () => {
    const settings = { ...DEFAULT_DIAGRAM_SETTINGS, sequenceNumbers: true, sequenceWrap: true }
    const configured = { ...existing, settings, source: 'sequenceDiagram\nA->>B: Hello' }
    const { result } = await openPane(configured)
    expect(result.current.draft.settings).toEqual(settings)
    expect(result.current.diagramKind).toBe('sequence')
    await renderPending()
    expect(updateDiagramById).not.toHaveBeenCalled()
    act(() => selectInWord({ ...existing, id: 'legacy-other' }))
    expect(result.current.draft.settings).toEqual(DEFAULT_DIAGRAM_SETTINGS)
    expect(result.current.diagramKind).toBe('flowchart')
  })

  it('stores settings on initial insertion and retains them on later source edits', async () => {
    const { result } = await openPane()
    const settings = { ...DEFAULT_DIAGRAM_SETTINGS, layout: 'elk' as const, imageQuality: 'standard' as const }
    act(() => result.current.applySettings('neutral', settings))
    await renderPending()
    await act(async () => { await result.current.insert() })
    expect(insertDiagramWithPayload).toHaveBeenCalledExactlyOnceWith(
      { svg: expect.any(String),
        draft: { source: DEFAULT_DIAGRAM, theme: 'neutral', size: 'medium', settings },
        requireEmptySelection: true },
    )
    expect(result.current.target?.settings).toEqual(settings)
    act(() => result.current.changeSource('flowchart LR\nUpdated'))
    await renderPending()
    expect(result.current.target?.settings).toEqual(settings)
  })

  it('does not rewrite unchanged settings or a legacy default diagram', async () => {
    const { result } = await openPane(existing)
    await renderPending()
    act(() => result.current.applySettings(existing.theme, { ...DEFAULT_DIAGRAM_SETTINGS }))
    await renderPending()
    expect(updateDiagramById).not.toHaveBeenCalled()
    expect(result.current.dirty).toBe(false)
  })

  it('confirms switching away from unapplied Word settings changes', async () => {
    const { result } = await openPane(existing)
    act(() => result.current.applySettings('dark', { ...DEFAULT_DIAGRAM_SETTINGS, sequenceWrap: true }))
    act(() => selectInWord(null))
    expect(result.current.pending).not.toBeNull()
    await renderPending()
    expect(updateDiagramById).not.toHaveBeenCalled()
  })

  it('retains newer settings while an earlier Word write finishes', async () => {
    const { result } = await openPane(existing)
    await renderPending()
    let finish: (format: 'png') => void = () => { throw new Error('Write has not started') }
    vi.mocked(updateDiagramById).mockImplementationOnce(() => new Promise(resolve => { finish = resolve }))
    const first = { ...DEFAULT_DIAGRAM_SETTINGS, fontSize: 18 as const }
    const latest = { ...first, imageQuality: 'high' as const }
    act(() => result.current.applySettings('dark', first))
    await renderPending()
    act(() => result.current.applySettings('dark', latest))
    await renderPending()
    expect(updateDiagramById).toHaveBeenCalledOnce()
    await act(async () => { finish('png') })
    expect(updateDiagramById).toHaveBeenCalledTimes(2)
    expect(updateDiagramById).toHaveBeenLastCalledWith(
      { svg: expect.any(String), existing: expect.objectContaining({ settings: first }),
        draft: { source: existing.source, theme: 'dark', size: 'medium', settings: latest },
        applySize: false },
    )
    expect(result.current.target?.settings).toEqual(latest)
    expect(result.current.draft.settings).toEqual(latest)
    expect(result.current.dirty).toBe(false)
  })
})
