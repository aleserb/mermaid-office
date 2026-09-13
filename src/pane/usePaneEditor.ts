import { useCallback, useEffect, useRef, useState } from 'react'
import { DEFAULT_DIAGRAM } from '../defaultDiagram'
import { normalizeMermaidError, type MermaidDiagnostic } from '../mermaid/diagnostics'
import { detectDiagramKind, renderMermaid } from '../mermaid/render'
import type { DiagramPayload, DiagramSize, DiagramTheme } from '../metadata/payload'
import { getDiagramSettings, sameDiagramSettings, type DiagramSettings } from '../metadata/diagramSettings'
import { getHostAdapter, type HostAdapter } from '../office/hostAdapter'
import { getPreferredSettings, getPreferredTheme, setPreferredSettings, setPreferredTheme } from '../preferences/diagramPreferences'
import type { DiagramSelectionWatcher } from '../office/selectionWatcher'
import { isLargeDiagram, requiresManualLargeDiagramUpdates } from './updateMode'

export const LIVE_UPDATE_DELAY = 600

interface Draft {
  source: string
  theme: DiagramTheme
  size: DiagramSize
  settings: DiagramSettings
}

interface RenderedDraft {
  draft: Draft
  svg: string
}

function newDraft(): Draft {
  return { source: DEFAULT_DIAGRAM, theme: getPreferredTheme(), size: 'medium', settings: getPreferredSettings() }
}

function sameDraft(
  left: Pick<DiagramPayload, 'source' | 'theme' | 'size' | 'settings'>,
  right: Pick<DiagramPayload, 'source' | 'theme' | 'size' | 'settings'>,
): boolean {
  return left.source === right.source && left.theme === right.theme && left.size === right.size
    && sameDiagramSettings(left.settings, right.settings)
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : 'Unable to update the Office diagram.'
}

export function usePaneEditor() {
  const [baseline, setBaseline] = useState(newDraft)
  const [draft, setDraft] = useState<Draft>(baseline)
  const [target, setTarget] = useState<DiagramPayload | null>(null)
  const [ready, setReady] = useState(false)
  const [hostName, setHostName] = useState<'Word' | 'Excel'>('Word')
  const [writing, setWriting] = useState(false)
  const [loadingSelection, setLoadingSelection] = useState(false)
  const [rendered, setRendered] = useState<RenderedDraft | null>(null)
  const [diagnostic, setDiagnostic] = useState<MermaidDiagnostic | null>(null)
  const [wordError, setWordError] = useState('')
  const [failedWrite, setFailedWrite] = useState<RenderedDraft | null>(null)
  const [pending, setPending] = useState<{ target: DiagramPayload | null } | null>(null)
  const [historyKey, setHistoryKey] = useState(0)
  const [settingsActive, setSettingsActive] = useState(false)
  const [manualUpdates, setManualUpdates] = useState(() => isLargeDiagram(baseline.source))
  const [updateRequested, setUpdateRequested] = useState<Draft | null>(null)
  const settingsPinned = useRef(false)
  const resumeAfterSettingsWrite = useRef(false)
  const manualLargeDiagramUpdates = useRef(true)
  const mounted = useRef(false)
  const busy = useRef(false)
  const draftRef = useRef(draft)
  const savedRef = useRef<Pick<DiagramPayload, 'source' | 'theme' | 'size' | 'settings'>>(baseline)
  const targetRef = useRef<DiagramPayload | null>(null)
  const lastSelectionId = useRef<string | null | undefined>(undefined)
  const selectionFromDocument = useRef(false)
  const watcher = useRef<DiagramSelectionWatcher | null>(null)
  const adapter = useRef<HostAdapter>(getHostAdapter())

  const activate = useCallback((payload: DiagramPayload | null) => {
    const next = payload
      ? { source: payload.source, theme: payload.theme, size: payload.size, settings: getDiagramSettings(payload.settings) }
      : newDraft()
    draftRef.current = next
    savedRef.current = next
    targetRef.current = payload
    setDraft(next)
    setBaseline(next)
    setTarget(payload)
    setRendered(null)
    setDiagnostic(null)
    setWordError('')
    setFailedWrite(null)
    setPending(null)
    setManualUpdates(manualLargeDiagramUpdates.current && isLargeDiagram(next.source))
    setUpdateRequested(null)
    setHistoryKey((key) => key + 1)
  }, [])

  const requestTarget = useCallback((payload: DiagramPayload | null) => {
    if (!sameDraft(draftRef.current, savedRef.current)) {
      setPending({ target: payload })
    } else {
      activate(payload)
    }
  }, [activate])

  const receiveSelection = useCallback((payload: DiagramPayload | null) => {
    if (settingsPinned.current) return
    const fromDocument = selectionFromDocument.current
    selectionFromDocument.current = false
    setLoadingSelection(false)
    const id = payload?.id ?? null

    // Picture replacement can clear the host's selection while focus stays in the
    // pane. A real document selection gesture, however, must switch the editor.
    if (!payload && targetRef.current && !fromDocument && document.hasFocus()) return
    if (id === lastSelectionId.current) return
    lastSelectionId.current = id
    if (id === (targetRef.current?.id ?? null)) {
      setPending(null)
      return
    }
    requestTarget(payload)
  }, [requestTarget])

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
        setHostName(adapter.current.appName)
        manualLargeDiagramUpdates.current = requiresManualLargeDiagramUpdates(
          Office.context.platform === undefined ? undefined : String(Office.context.platform),
        )
        watcher.current = adapter.current.watchSelectedDiagram(
          (selected) => {
            if (!active) return
            receiveSelection(selected)
            setReady(true)
          },
          (error) => {
            if (!active) return
            setWordError(error.message)
            if (!settingsPinned.current) setLoadingSelection(true)
            setReady(true)
          },
          {
            isPaused: () => busy.current || settingsPinned.current,
            onSelectionChange: () => {
              if (!active) return
              selectionFromDocument.current ||= !document.hasFocus()
              if (settingsPinned.current) return
              setLoadingSelection(true)
            },
          },
        )
      } catch (error) {
        if (active) {
          setWordError(errorMessage(error))
          setReady(true)
        }
      }
    }
    void initialize()
    return () => {
      active = false
      mounted.current = false
      watcher.current?.()
      watcher.current = null
    }
  }, [receiveSelection])

  useEffect(() => {
    if (!ready) return
    let active = true
    const timer = window.setTimeout(async () => {
      try {
        const svg = await renderMermaid(draft.source, draft.theme, draft.settings)
        if (active) {
          if (manualLargeDiagramUpdates.current && isLargeDiagram(draft.source, svg)) {
            setManualUpdates(true)
          }
          setRendered({ draft, svg })
        }
      } catch (error) {
        if (active) setDiagnostic(normalizeMermaidError(error, draft.source))
      }
    }, LIVE_UPDATE_DELAY)
    return () => {
      active = false
      window.clearTimeout(timer)
    }
  }, [draft, ready])

  useEffect(() => {
    if (
      !ready || !target || !rendered || writing || busy.current || settingsActive || loadingSelection || pending ||
      !sameDraft(rendered.draft, draft) || sameDraft(draft, target) || failedWrite === rendered ||
      (manualUpdates && updateRequested !== draft)
    ) {
      return
    }

    // Only one host write may be in flight. Edits made during it remain in draft
    // for the next automatic pass or explicit Update, never replaced by selection.
    busy.current = true
    setUpdateRequested(null)
    setWriting(true)
    setWordError('')
    void adapter.current.updateDiagramById(
      rendered.svg, target, draft.source, draft.theme, draft.size, draft.size !== target.size,
      undefined, draft.settings,
    ).then((format) => {
      const saved = { ...target, ...draft, format }
      savedRef.current = saved
      targetRef.current = saved
      if (mounted.current) {
        setTarget(saved)
        setFailedWrite(null)
      }
    }).catch((error: unknown) => {
      console.error('Unable to update the pane diagram.', error)
      if (mounted.current) {
        setWordError(errorMessage(error))
        setFailedWrite(rendered)
      }
    }).finally(() => {
      busy.current = false
      if (mounted.current) {
        setWriting(false)
        if (resumeAfterSettingsWrite.current) {
          resumeAfterSettingsWrite.current = false
          settingsPinned.current = false
        }
        watcher.current?.refresh()
      }
    })
  }, [draft, failedWrite, loadingSelection, manualUpdates, pending, ready, rendered, settingsActive, target, updateRequested, writing])

  useEffect(() => {
    if (!settingsActive && resumeAfterSettingsWrite.current && !writing &&
      (diagnostic || (failedWrite && failedWrite === rendered) || !target || sameDraft(draft, target) ||
        (manualUpdates && !updateRequested))) {
      resumeAfterSettingsWrite.current = false
      settingsPinned.current = false
      watcher.current?.refresh()
    }
  }, [diagnostic, draft, failedWrite, manualUpdates, rendered, settingsActive, target, updateRequested, writing])

  const beginSettings = () => {
    settingsPinned.current = true
    resumeAfterSettingsWrite.current = false
    setSettingsActive(true)
  }

  const endSettings = (applied: boolean) => {
    setSettingsActive(false)
    if (applied && targetRef.current && !sameDraft(draftRef.current, targetRef.current)) {
      resumeAfterSettingsWrite.current = true
    } else {
      settingsPinned.current = false
      watcher.current?.refresh()
    }
  }

  const changeDraft = (next: Draft) => {
    if (sameDraft(next, draftRef.current)) return
    draftRef.current = next
    setDraft(next)
    setRendered(null)
    setDiagnostic(null)
    setUpdateRequested(null)
    if (manualLargeDiagramUpdates.current && isLargeDiagram(next.source)) {
      setManualUpdates(true)
    }
  }

  const changeSource = (source: string) => changeDraft({ ...draftRef.current, source })

  const changeTheme = (theme: DiagramTheme) => {
    changeDraft({ ...draftRef.current, theme })
    setPreferredTheme(theme)
  }

  const applySettings = (theme: DiagramTheme, settings: DiagramSettings) => {
    changeDraft({ ...draftRef.current, theme, settings: { ...settings } })
    if (targetRef.current) setUpdateRequested(draftRef.current)
    setPreferredTheme(theme)
    setPreferredSettings(settings)
  }

  const insert = async () => {
    if (busy.current) return
    if (!rendered || !sameDraft(rendered.draft, draft) || diagnostic || target) {
      setWordError('Wait for a valid new diagram before inserting.')
      return
    }
    busy.current = true
    setWriting(true)
    setWordError('')
    try {
      if (await adapter.current.getSelectedDiagram()) {
        throw new Error(adapter.current.insertionLocationError)
      }
      const inserted = await adapter.current.insertDiagramWithPayload(
        rendered.svg, draft.source, draft.theme, draft.size, undefined, { requireEmptySelection: true },
        draft.settings,
      )
      savedRef.current = inserted
      targetRef.current = inserted
      lastSelectionId.current = inserted.id
      if (mounted.current) setTarget(inserted)
    } catch (error) {
      if (mounted.current) setWordError(errorMessage(error))
    } finally {
      busy.current = false
      if (mounted.current) {
        setWriting(false)
        watcher.current?.refresh()
      }
    }
  }

  const canUpdate = ready && !!target && !sameDraft(draft, target) &&
    Boolean(rendered && sameDraft(rendered.draft, draft)) && !diagnostic &&
    !writing && !loadingSelection && !settingsActive && !pending

  const update = () => {
    if (!canUpdate || busy.current) {
      setWordError('Wait for a valid diagram and finish the current action before updating.')
      return
    }
    setFailedWrite(null)
    setUpdateRequested(draft)
  }

  return {
    draft, target, ready, hostName, writing, loadingSelection, diagnostic, wordError, pending, historyKey, settingsActive,
    manualUpdates, canUpdate, update,
    dirty: !sameDraft(draft, target ?? baseline),
    rendering: !rendered || !sameDraft(rendered.draft, draft),
    canInsert: ready && !target && Boolean(rendered && sameDraft(rendered.draft, draft)) &&
      !diagnostic && !writing && !loadingSelection && !settingsActive,
    canRetry: Boolean(failedWrite && rendered === failedWrite),
    changeSource, changeTheme, applySettings, insert, beginSettings, endSettings,
    diagramKind: detectDiagramKind(draft.source),
    keepEditing: () => setPending(null),
    discardAndSwitch: () => { if (pending && !busy.current) activate(pending.target) },
    retry: () => setFailedWrite(null),
  }
}
