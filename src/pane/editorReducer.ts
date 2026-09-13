import type { MermaidDiagnostic } from '../mermaid/diagnostics'
import { getDiagramSettings, sameDiagramSettings } from '../metadata/diagramSettings'
import type { DiagramPayload } from '../metadata/payload'
import type { DiagramDraft } from '../office/diagramRequests'
import { isDocumentSelectionGesture, type SelectionOrigin } from '../office/selectionWatcher'
import { isLargeDiagram } from './updateMode'

export interface RenderedDraft {
  draft: DiagramDraft
  svg: string
}

export interface EditorState {
  baseline: DiagramDraft
  draft: DiagramDraft
  target: DiagramPayload | null
  ready: boolean
  hostName: 'Word' | 'Excel'
  writing: boolean
  loadingSelection: boolean
  rendered: RenderedDraft | null
  diagnostic: MermaidDiagnostic | null
  hostError: string
  failedWrite: RenderedDraft | null
  pending: { target: DiagramPayload | null } | null
  historyKey: number
  settingsActive: boolean
  settingsPinned: boolean
  resumeAfterSettingsWrite: boolean
  manualLargeDiagramUpdates: boolean
  manualUpdates: boolean
  updateRequested: DiagramDraft | null
  selectionId: string | null | undefined
}

export function sameDraft(
  left: Pick<DiagramPayload, 'source' | 'theme' | 'size' | 'settings'>,
  right: Pick<DiagramPayload, 'source' | 'theme' | 'size' | 'settings'>,
): boolean {
  return left.source === right.source && left.theme === right.theme && left.size === right.size
    && sameDiagramSettings(left.settings, right.settings)
}

export function createEditorState(draft: DiagramDraft): EditorState {
  return {
    baseline: draft, draft, target: null, ready: false, hostName: 'Word',
    writing: false, loadingSelection: false, rendered: null, diagnostic: null,
    hostError: '', failedWrite: null, pending: null, historyKey: 0,
    settingsActive: false, settingsPinned: false, resumeAfterSettingsWrite: false,
    manualLargeDiagramUpdates: true, manualUpdates: isLargeDiagram(draft.source),
    updateRequested: null, selectionId: undefined,
  }
}

export type EditorAction =
  | { type: 'initialized'; hostName: 'Word' | 'Excel'; manualLargeDiagramUpdates: boolean }
  | { type: 'selection-started' }
  | { type: 'selection'; payload: DiagramPayload | null; origin: SelectionOrigin; draft: DiagramDraft }
  | { type: 'selection-failed'; message: string }
  | { type: 'initialization-failed'; message: string }
  | { type: 'change-draft'; draft: DiagramDraft }
  | { type: 'rendered'; result: RenderedDraft }
  | { type: 'render-failed'; draft: DiagramDraft; diagnostic: MermaidDiagnostic }
  | { type: 'write-started' }
  | { type: 'write-succeeded'; payload: DiagramPayload; inserted?: boolean }
  | { type: 'write-failed'; message: string; rendered?: RenderedDraft }
  | { type: 'write-finished' }
  | { type: 'host-error'; message: string }
  | { type: 'begin-settings' }
  | { type: 'end-settings'; applied: boolean }
  | { type: 'unpin-settings' }
  | { type: 'request-update'; retry?: boolean }
  | { type: 'retry' }
  | { type: 'keep-editing' }
  | { type: 'discard-and-switch'; draft: DiagramDraft }

export function ignoresSelection(state: EditorState, payload: DiagramPayload | null, origin: SelectionOrigin): boolean {
  return state.settingsPinned || (!payload && !!state.target && !isDocumentSelectionGesture(origin))
}

export function selectionDraft(payload: DiagramPayload | null, emptyDraft: DiagramDraft): DiagramDraft {
  return payload
    ? { source: payload.source, theme: payload.theme, size: payload.size, settings: getDiagramSettings(payload.settings) }
    : emptyDraft
}

function activate(state: EditorState, payload: DiagramPayload | null, draft: DiagramDraft): EditorState {
  return {
    ...state, draft, baseline: draft, target: payload, rendered: null, diagnostic: null,
    hostError: '', failedWrite: null, pending: null, updateRequested: null,
    historyKey: state.historyKey + 1,
    manualUpdates: state.manualLargeDiagramUpdates && isLargeDiagram(draft.source),
  }
}

export function editorReducer(state: EditorState, action: EditorAction): EditorState {
  switch (action.type) {
    case 'initialized':
      return {
        ...state, hostName: action.hostName, manualLargeDiagramUpdates: action.manualLargeDiagramUpdates,
        manualUpdates: action.manualLargeDiagramUpdates && isLargeDiagram(state.draft.source),
      }
    case 'selection-started':
      return state.settingsPinned ? state : { ...state, loadingSelection: true }
    case 'selection': {
      const next = { ...state, ready: true, loadingSelection: false }
      if (ignoresSelection(state, action.payload, action.origin)) return state.settingsPinned ? state : next
      const id = action.payload?.id ?? null
      if (id === state.selectionId) return next
      next.selectionId = id
      if (id === (state.target?.id ?? null)) return { ...next, pending: null }
      return sameDraft(state.draft, state.target ?? state.baseline)
        ? activate(next, action.payload, action.draft)
        : { ...next, pending: { target: action.payload } }
    }
    case 'selection-failed':
      return { ...state, ready: true, hostError: action.message, loadingSelection: !state.settingsPinned }
    case 'initialization-failed':
      return { ...state, ready: true, hostError: action.message }
    case 'change-draft':
      if (sameDraft(state.draft, action.draft)) return state
      return {
        ...state, draft: action.draft, rendered: null, diagnostic: null, updateRequested: null,
        manualUpdates: state.manualUpdates || (state.manualLargeDiagramUpdates && isLargeDiagram(action.draft.source)),
      }
    case 'rendered':
      if (action.result.draft !== state.draft) return state
      return {
        ...state, rendered: action.result, diagnostic: null,
        manualUpdates: state.manualUpdates ||
          (state.manualLargeDiagramUpdates && isLargeDiagram(state.draft.source, action.result.svg)),
      }
    case 'render-failed':
      return action.draft !== state.draft ? state : { ...state, diagnostic: action.diagnostic }
    case 'write-started':
      return { ...state, writing: true, hostError: '', updateRequested: null }
    case 'write-succeeded':
      return {
        ...state, target: action.payload, failedWrite: null,
        selectionId: action.inserted ? action.payload.id : state.selectionId,
      }
    case 'write-failed':
      return { ...state, hostError: action.message, failedWrite: action.rendered ?? state.failedWrite }
    case 'write-finished':
      return {
        ...state, writing: false,
        settingsPinned: state.resumeAfterSettingsWrite ? false : state.settingsPinned,
        resumeAfterSettingsWrite: false,
      }
    case 'host-error':
      return { ...state, hostError: action.message }
    case 'begin-settings':
      return { ...state, settingsActive: true, settingsPinned: true, resumeAfterSettingsWrite: false }
    case 'end-settings': {
      const resume = action.applied && !!state.target && !sameDraft(state.draft, state.target)
      return { ...state, settingsActive: false, settingsPinned: resume, resumeAfterSettingsWrite: resume }
    }
    case 'unpin-settings':
      return { ...state, settingsPinned: false, resumeAfterSettingsWrite: false }
    case 'request-update':
      return {
        ...state, updateRequested: state.draft,
        failedWrite: action.retry ? null : state.failedWrite,
      }
    case 'retry':
      return { ...state, failedWrite: null }
    case 'keep-editing':
      return { ...state, pending: null }
    case 'discard-and-switch':
      return state.pending && !state.writing ? activate(state, state.pending.target, action.draft) : state
  }
}
