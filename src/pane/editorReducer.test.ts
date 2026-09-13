import { describe, expect, it } from 'vitest'
import { DEFAULT_DIAGRAM_SETTINGS } from '../metadata/diagramSettings'
import { createDiagramPayload } from '../metadata/payload'
import type { DiagramDraft } from '../office/diagramRequests'
import type { SelectionOrigin } from '../office/selectionWatcher'
import { createEditorState, editorReducer, ignoresSelection, selectionDraft, type EditorState } from './editorReducer'

const emptyDraft: DiagramDraft = {
  source: 'flowchart LR\nNew', theme: 'default', size: 'medium', settings: DEFAULT_DIAGRAM_SETTINGS,
}
const first = createDiagramPayload('flowchart LR\nFirst', 'png')
const second = createDiagramPayload('flowchart LR\nSecond', 'png')

function select(state: EditorState, payload: typeof first | null, origin: SelectionOrigin = 'office-event') {
  return editorReducer(state, { type: 'selection', payload, origin, draft: selectionDraft(payload, emptyDraft) })
}

function editingFirst() {
  return select(createEditorState(emptyDraft), first, 'initial')
}

describe('pure editor transitions', () => {
  it('ignores an unknown null without recording it as the selection', () => {
    const state = editingFirst()
    expect(ignoresSelection(state, null, 'background-poll')).toBe(true)
    const ignored = select(state, null, 'background-poll')
    expect(ignored.target).toBe(first)
    expect(ignored.draft).toBe(state.draft)
    expect(ignored.selectionId).toBe(first.id)
    expect(ignored.historyKey).toBe(state.historyKey)
    const deselected = select(ignored, null, 'pane-focus')
    expect(deselected.target).toBeNull()
    expect(deselected.draft).toBe(emptyDraft)
    expect(deselected.historyKey).toBe(state.historyKey + 1)
  })

  it('does not mutate state or reset dirty drafts on repeated selection snapshots', () => {
    const state = editingFirst()
    Object.freeze(state)
    const draft = { ...state.draft, source: 'flowchart LR\nUnsaved' }
    const edited = editorReducer(state, { type: 'change-draft', draft })
    const repeated = select(edited, first)
    expect(repeated.draft).toBe(draft)
    expect(repeated.historyKey).toBe(state.historyKey)
    expect(repeated.pending).toBeNull()
    expect(state.draft.source).toBe(first.source)
  })

  it('keeps only the latest pending selection and resets history on confirmed switching', () => {
    const original = editingFirst()
    const edited = editorReducer(original, {
      type: 'change-draft', draft: { ...original.draft, source: 'flowchart LR\nUnsaved' },
    })
    const pending = select(select(edited, second), null)
    expect(pending.target).toBe(first)
    expect(pending.pending).toEqual({ target: null })
    const switched = editorReducer(pending, { type: 'discard-and-switch', draft: emptyDraft })
    expect(switched.target).toBeNull()
    expect(switched.draft).toBe(emptyDraft)
    expect(switched.historyKey).toBe(original.historyKey + 1)
  })

  it('cancels confirmation when the current diagram is reselected', () => {
    const state = editingFirst()
    const edited = editorReducer(state, { type: 'change-draft', draft: { ...state.draft, source: 'unsaved' } })
    const reselected = select(select(edited, second), first)
    expect(reselected.pending).toBeNull()
    expect(reselected.draft).toBe(edited.draft)
  })

  it('ignores stale rendering success and failure after an edit or selection switch', () => {
    const state = editingFirst()
    const switched = select(state, second)
    expect(editorReducer(switched, { type: 'rendered', result: { draft: state.draft, svg: '<svg/>' } })).toBe(switched)
    expect(editorReducer(switched, {
      type: 'render-failed', draft: state.draft, diagnostic: { message: 'Old error', line: 1, column: 1, endColumn: 2 },
    })).toBe(switched)
  })

  it('uses the saved write revision without replacing edits made during the write', () => {
    const state = editingFirst()
    const writing = editorReducer(state, { type: 'write-started' })
    const draft = { ...state.draft, source: 'flowchart LR\nTypedDuringWrite' }
    const edited = editorReducer(writing, { type: 'change-draft', draft })
    const saved = { ...first, source: 'flowchart LR\nSaved' }
    const completed = editorReducer(edited, { type: 'write-succeeded', payload: saved })
    expect(completed.target).toBe(saved)
    expect(completed.draft).toBe(draft)
    expect(completed.historyKey).toBe(state.historyKey)
  })

  it('pins applied settings until the write finishes', () => {
    const state = editingFirst()
    const pinned = editorReducer(state, { type: 'begin-settings' })
    expect(select(pinned, second)).toBe(pinned)
    const edited = editorReducer(pinned, { type: 'change-draft', draft: { ...pinned.draft, theme: 'dark' } })
    const applied = editorReducer(edited, { type: 'end-settings', applied: true })
    expect(applied.settingsPinned).toBe(true)
    expect(applied.settingsActive).toBe(false)
    expect(editorReducer(applied, { type: 'write-finished' }).settingsPinned).toBe(false)
  })

  it('retains failure identity until retry or a new render', () => {
    const state = editingFirst()
    const rendered = { draft: state.draft, svg: '<svg/>' }
    const failed = editorReducer(state, { type: 'write-failed', message: 'Unavailable', rendered })
    expect(failed.failedWrite).toBe(rendered)
    expect(failed.hostError).toBe('Unavailable')
    expect(editorReducer(failed, { type: 'retry' }).failedWrite).toBeNull()
  })
})
