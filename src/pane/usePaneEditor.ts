import { useCallback, useEffect, useReducer, useRef } from 'react'
import { DEFAULT_DIAGRAM } from '../defaultDiagram'
import { detectDiagramKind } from '../mermaid/render'
import type { DiagramTheme } from '../metadata/payload'
import type { DiagramSettings } from '../metadata/diagramSettings'
import type { DiagramDraft } from '../office/diagramRequests'
import { getHostAdapter, type HostAdapter } from '../office/hostAdapter'
import { getPreferredSettings, getPreferredTheme, setPreferredSettings, setPreferredTheme } from '../preferences/diagramPreferences'
import type { DiagramSelectionWatcher, SelectionReceiver } from '../office/selectionWatcher'
import { requiresManualLargeDiagramUpdates } from './updateMode'
import { createEditorState, editorReducer, ignoresSelection, sameDraft, selectionDraft, type EditorAction } from './editorReducer'
import { useDebouncedDiagramRender } from './useDebouncedDiagramRender'

export { LIVE_UPDATE_DELAY } from './useDebouncedDiagramRender'

function newDraft(): DiagramDraft {
  return { source: DEFAULT_DIAGRAM, theme: getPreferredTheme(), size: 'medium', settings: getPreferredSettings() }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : 'Unable to update the Office diagram.'
}

export function usePaneEditor() {
  const [state, reduce] = useReducer(editorReducer, undefined, () => createEditorState(newDraft()))
  // Office callbacks and a settings Apply/Close pair can run before React commits.
  // Keep their view current by applying the same pure transition synchronously.
  const current = useRef(state)
  const dispatch = useCallback((action: EditorAction) => {
    current.current = editorReducer(current.current, action)
    reduce(action)
  }, [])
  const {
    baseline, draft, target, ready, hostName, writing, loadingSelection, rendered, diagnostic,
    hostError, failedWrite, pending, historyKey, settingsActive, manualUpdates, updateRequested,
    resumeAfterSettingsWrite,
  } = state
  const mounted = useRef(false)
  const busy = useRef(false)
  const watcher = useRef<DiagramSelectionWatcher | null>(null)
  const adapter = useRef<HostAdapter>(getHostAdapter())

  const receiveSelection = useCallback<SelectionReceiver>((payload, origin) => {
    const accepted = !ignoresSelection(current.current, payload, origin)
    dispatch({ type: 'selection', payload, origin, draft: selectionDraft(payload, newDraft()) })
    return accepted
  }, [dispatch])

  useEffect(() => {
    mounted.current = true
    let active = true
    const initialize = async () => {
      try {
        if (typeof Office === 'undefined') {
          throw new Error('Open Mermaid inside Microsoft Word or Excel to insert or edit diagrams.')
        }
        const info = await Office.onReady()
        if (!active) return
        if (!Office.context?.document) {
          throw new Error('Open Mermaid inside Microsoft Word or Excel to follow document selection.')
        }
        adapter.current = getHostAdapter(info.host ?? Office.context.host)
        dispatch({
          type: 'initialized',
          hostName: adapter.current.appName,
          manualLargeDiagramUpdates: requiresManualLargeDiagramUpdates(
            Office.context.platform === undefined ? undefined : String(Office.context.platform),
          ),
        })
        watcher.current = adapter.current.watchSelectedDiagram(
          (selected, origin) => active ? receiveSelection(selected, origin) : false,
          (error) => {
            if (active) dispatch({ type: 'selection-failed', message: error.message })
          },
          {
            isPaused: () => busy.current || current.current.settingsPinned,
            onSelectionChange: () => {
              if (active) dispatch({ type: 'selection-started' })
            },
          },
        )
      } catch (error) {
        if (active) dispatch({ type: 'initialization-failed', message: errorMessage(error) })
      }
    }
    void initialize()
    return () => {
      active = false
      mounted.current = false
      watcher.current?.()
      watcher.current = null
    }
  }, [dispatch, receiveSelection])

  useDebouncedDiagramRender(draft, ready, dispatch)

  useEffect(() => {
    if (
      !ready || !target || !rendered || writing || busy.current || settingsActive || loadingSelection || pending ||
      !sameDraft(rendered.draft, draft) || sameDraft(draft, target) || failedWrite === rendered ||
      (manualUpdates && updateRequested !== draft)
    ) return

    busy.current = true
    dispatch({ type: 'write-started' })
    void adapter.current.updateDiagramById({
      svg: rendered.svg, existing: target, draft, applySize: draft.size !== target.size,
    }).then((format) => {
      if (mounted.current) dispatch({ type: 'write-succeeded', payload: { ...target, ...draft, format } })
    }).catch((error: unknown) => {
      console.error('Unable to update the pane diagram.', error)
      if (mounted.current) dispatch({ type: 'write-failed', message: errorMessage(error), rendered })
    }).finally(() => {
      busy.current = false
      if (mounted.current) {
        dispatch({ type: 'write-finished' })
        watcher.current?.refresh()
      }
    })
  }, [dispatch, draft, failedWrite, loadingSelection, manualUpdates, pending, ready, rendered, settingsActive, target, updateRequested, writing])

  useEffect(() => {
    if (!settingsActive && resumeAfterSettingsWrite && !writing &&
      (diagnostic || (failedWrite && failedWrite === rendered) || !target || sameDraft(draft, target) ||
        (manualUpdates && !updateRequested))) {
      dispatch({ type: 'unpin-settings' })
      watcher.current?.refresh()
    }
  }, [dispatch, diagnostic, draft, failedWrite, manualUpdates, rendered, resumeAfterSettingsWrite, settingsActive, target, updateRequested, writing])

  const beginSettings = () => dispatch({ type: 'begin-settings' })
  const endSettings = (applied: boolean) => {
    dispatch({ type: 'end-settings', applied })
    if (!current.current.settingsPinned) watcher.current?.refresh()
  }
  const changeSource = (source: string) => dispatch({ type: 'change-draft', draft: { ...current.current.draft, source } })
  const changeTheme = (theme: DiagramTheme) => {
    dispatch({ type: 'change-draft', draft: { ...current.current.draft, theme } })
    setPreferredTheme(theme)
  }
  const applySettings = (theme: DiagramTheme, settings: DiagramSettings) => {
    dispatch({ type: 'change-draft', draft: { ...current.current.draft, theme, settings: { ...settings } } })
    if (current.current.target) dispatch({ type: 'request-update' })
    setPreferredTheme(theme)
    setPreferredSettings(settings)
  }

  const insert = async () => {
    if (busy.current) return
    if (!rendered || !sameDraft(rendered.draft, draft) || diagnostic || target) {
      dispatch({ type: 'host-error', message: 'Wait for a valid new diagram before inserting.' })
      return
    }
    busy.current = true
    dispatch({ type: 'write-started' })
    try {
      if (await adapter.current.getSelectedDiagram()) throw new Error(adapter.current.insertionLocationError)
      const inserted = await adapter.current.insertDiagramWithPayload({
        svg: rendered.svg, draft, requireEmptySelection: true,
      })
      if (mounted.current) dispatch({ type: 'write-succeeded', payload: inserted, inserted: true })
    } catch (error) {
      if (mounted.current) dispatch({ type: 'write-failed', message: errorMessage(error) })
    } finally {
      busy.current = false
      if (mounted.current) {
        dispatch({ type: 'write-finished' })
        watcher.current?.refresh()
      }
    }
  }

  const canUpdate = ready && !!target && !sameDraft(draft, target) &&
    Boolean(rendered && sameDraft(rendered.draft, draft)) && !diagnostic &&
    !writing && !loadingSelection && !settingsActive && !pending
  const update = () => {
    if (!canUpdate || busy.current) {
      dispatch({ type: 'host-error', message: 'Wait for a valid diagram and finish the current action before updating.' })
      return
    }
    dispatch({ type: 'request-update', retry: true })
  }

  return {
    draft, target, ready, hostName, writing, loadingSelection, diagnostic, hostError, pending, historyKey, settingsActive,
    manualUpdates, canUpdate, update,
    dirty: !sameDraft(draft, target ?? baseline),
    rendering: !rendered || !sameDraft(rendered.draft, draft),
    canInsert: ready && !target && Boolean(rendered && sameDraft(rendered.draft, draft)) &&
      !diagnostic && !writing && !loadingSelection && !settingsActive,
    canRetry: Boolean(failedWrite && rendered === failedWrite),
    changeSource, changeTheme, applySettings, insert, beginSettings, endSettings,
    diagramKind: detectDiagramKind(draft.source),
    keepEditing: () => dispatch({ type: 'keep-editing' }),
    discardAndSwitch: () => {
      if (!busy.current) dispatch({
        type: 'discard-and-switch', draft: selectionDraft(current.current.pending?.target ?? null, newDraft()),
      })
    },
    retry: () => dispatch({ type: 'retry' }),
  }
}
