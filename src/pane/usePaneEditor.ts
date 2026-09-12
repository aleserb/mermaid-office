import { useCallback, useEffect, useRef, useState } from 'react'
import { DEFAULT_DIAGRAM } from '../defaultDiagram'
import { normalizeMermaidError, type MermaidDiagnostic } from '../mermaid/diagnostics'
import { renderMermaid } from '../mermaid/render'
import type { DiagramPayload, DiagramSize, DiagramTheme } from '../metadata/payload'
import { getPreferredTheme, setPreferredTheme } from '../preferences/diagramPreferences'
import { insertDiagramWithPayload, updateDiagramById } from '../word/insertDiagram'
import { getSelectedDiagram, watchSelectedDiagram, type DiagramSelectionWatcher } from '../word/selection'

export const LIVE_UPDATE_DELAY = 600

interface Draft {
  source: string
  theme: DiagramTheme
  size: DiagramSize
}

interface RenderedDraft {
  draft: Draft
  svg: string
}

function newDraft(): Draft {
  return { source: DEFAULT_DIAGRAM, theme: getPreferredTheme(), size: 'medium' }
}

function sameDraft(left: Draft, right: Draft): boolean {
  return left.source === right.source && left.theme === right.theme && left.size === right.size
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : 'Unable to update the Word diagram.'
}

export function usePaneEditor() {
  const [baseline, setBaseline] = useState(newDraft)
  const [draft, setDraft] = useState<Draft>(baseline)
  const [target, setTarget] = useState<DiagramPayload | null>(null)
  const [ready, setReady] = useState(false)
  const [writing, setWriting] = useState(false)
  const [loadingSelection, setLoadingSelection] = useState(false)
  const [rendered, setRendered] = useState<RenderedDraft | null>(null)
  const [diagnostic, setDiagnostic] = useState<MermaidDiagnostic | null>(null)
  const [wordError, setWordError] = useState('')
  const [failedWrite, setFailedWrite] = useState<RenderedDraft | null>(null)
  const [pending, setPending] = useState<{ target: DiagramPayload | null } | null>(null)
  const [historyKey, setHistoryKey] = useState(0)
  const mounted = useRef(false)
  const busy = useRef(false)
  const draftRef = useRef(draft)
  const savedRef = useRef<Draft>(baseline)
  const targetRef = useRef<DiagramPayload | null>(null)
  const lastSelectionId = useRef<string | null | undefined>(undefined)
  const selectionFromDocument = useRef(false)
  const watcher = useRef<DiagramSelectionWatcher | null>(null)

  const activate = useCallback((payload: DiagramPayload | null) => {
    const next = payload
      ? { source: payload.source, theme: payload.theme, size: payload.size }
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
    const fromDocument = selectionFromDocument.current
    selectionFromDocument.current = false
    setLoadingSelection(false)
    const id = payload?.id ?? null

    // Picture replacement can clear Word's selection while focus stays in the
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
          throw new Error('Open Mermaid pane inside Microsoft Word to insert or edit diagrams.')
        }
        await Office.onReady()
        if (!active) return
        if (!Office.context?.document) {
          throw new Error('Open Mermaid pane inside Microsoft Word to follow document selection.')
        }
        watcher.current = watchSelectedDiagram(
          (selected) => {
            if (!active) return
            receiveSelection(selected)
            setReady(true)
          },
          (error) => {
            if (!active) return
            setWordError(error.message)
            setLoadingSelection(true)
            setReady(true)
          },
          {
            isPaused: () => busy.current,
            onSelectionChange: () => {
              if (!active) return
              selectionFromDocument.current ||= !document.hasFocus()
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
        const svg = await renderMermaid(draft.source, draft.theme)
        if (active) setRendered({ draft, svg })
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
      !ready || !target || !rendered || writing || busy.current || loadingSelection || pending ||
      !sameDraft(rendered.draft, draft) || sameDraft(draft, target) || failedWrite === rendered
    ) {
      return
    }

    // Only one Word write may be in flight. Edits made during it remain in draft
    // and are written on the next pass, never replaced by a selection notification.
    busy.current = true
    setWriting(true)
    setWordError('')
    void updateDiagramById(
      rendered.svg, target, draft.source, draft.theme, draft.size, draft.size !== target.size,
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
        watcher.current?.refresh()
      }
    })
  }, [draft, failedWrite, loadingSelection, pending, ready, rendered, target, writing])

  const changeDraft = (next: Draft) => {
    draftRef.current = next
    setDraft(next)
    setRendered(null)
    setDiagnostic(null)
  }

  const changeSource = (source: string) => changeDraft({ ...draftRef.current, source })

  const changeTheme = (theme: DiagramTheme) => {
    changeDraft({ ...draftRef.current, theme })
    setPreferredTheme(theme)
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
      if (await getSelectedDiagram()) {
        throw new Error('Place the cursor on a blank line before inserting a new diagram.')
      }
      const inserted = await insertDiagramWithPayload(
        rendered.svg, draft.source, draft.theme, draft.size, undefined, { requireEmptySelection: true },
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

  return {
    draft, target, ready, writing, loadingSelection, diagnostic, wordError, pending, historyKey,
    dirty: !sameDraft(draft, target ?? baseline),
    rendering: !rendered || !sameDraft(rendered.draft, draft),
    canInsert: ready && !target && Boolean(rendered && sameDraft(rendered.draft, draft)) &&
      !diagnostic && !writing && !loadingSelection,
    canRetry: Boolean(failedWrite && rendered === failedWrite),
    changeSource, changeTheme, insert,
    keepEditing: () => setPending(null),
    discardAndSwitch: () => { if (pending && !busy.current) activate(pending.target) },
    retry: () => setFailedWrite(null),
  }
}
