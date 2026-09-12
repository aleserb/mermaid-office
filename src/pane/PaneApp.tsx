import { useEffect, useRef, useState } from 'react'
import {
  Button,
  Caption1,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  FluentProvider,
  MessageBar,
  MessageBarBody,
  Spinner,
  webLightTheme,
} from '@fluentui/react-components'
import { MermaidEditor } from '../components/MermaidEditor'
import { SyntaxHelpLink } from '../components/SyntaxHelpLink'
import { DiagramSettingsDialog } from '../components/DiagramSettingsDialog'
import { usePaneEditor } from './usePaneEditor'
import { openSettingsWindow } from '../settings/openSettingsWindow'
import './pane.css'

export function PaneApp() {
  const editor = usePaneEditor()
  const settingsButton = useRef<HTMLButtonElement>(null)
  const restoreSettingsFocus = useRef(false)
  const [settingsHistoryKey, setSettingsHistoryKey] = useState<number | null>(null)
  const [settingsError, setSettingsError] = useState('')
  const settingsWindow = useRef<ReturnType<typeof openSettingsWindow> | null>(null)
  useEffect(() => () => settingsWindow.current?.dispose(), [])
  if (settingsHistoryKey !== null && (settingsHistoryKey !== editor.historyKey || editor.pending)) {
    setSettingsHistoryKey(null)
  }
  useEffect(() => {
    if (settingsHistoryKey === null && restoreSettingsFocus.current) {
      restoreSettingsFocus.current = false
      settingsButton.current?.focus()
    }
  }, [settingsHistoryKey])
  const closeSettings = () => {
    restoreSettingsFocus.current = true
    setSettingsHistoryKey(null)
  }
  const busy = !editor.ready || editor.writing || editor.loadingSelection || editor.settingsActive
  const openSettings = () => {
    setSettingsError('')
    if (typeof Office === 'undefined' || !Office.context?.document || !Office.context.ui?.displayDialogAsync) {
      setSettingsHistoryKey(editor.historyKey)
      return
    }
    editor.beginSettings()
    let applied = false
    settingsWindow.current = openSettingsWindow({
      theme: editor.draft.theme, settings: editor.draft.settings, diagramKind: editor.diagramKind,
    }, {
      onApply: (theme, settings) => {
        editor.applySettings(theme, settings)
        applied = true
      },
      onClose: () => {
        settingsWindow.current = null
        editor.endSettings(applied)
      },
      onError: setSettingsError,
    })
  }
  const status = editor.writing
    ? 'Writing diagram to Word...'
    : editor.target && editor.dirty
      ? editor.manualUpdates
        ? 'Changes not yet written to Word. Click Update to apply.'
        : 'Changes not yet written to Word.'
      : ''

  return (
    <FluentProvider theme={webLightTheme} style={{ colorScheme: 'light' }}>
      <main className="pane-shell" inert={editor.settingsActive}>
        <header className="pane-header">
          <div className="pane-actions">
            {!editor.target && (
              <Button appearance="primary" size="small" disabled={!editor.canInsert} onClick={() => void editor.insert()}>
                Insert
              </Button>
            )}
            {editor.target && editor.manualUpdates && (
              <Button appearance="primary" size="small" disabled={!editor.canUpdate} onClick={editor.update}
                title="Large diagrams update only when requested">
                {editor.canRetry ? 'Retry update' : 'Update'}
              </Button>
            )}
            {editor.canRetry && !editor.manualUpdates && (
              <Button size="small" disabled={busy} onClick={editor.retry}>Retry update</Button>
            )}
          </div>
          <div className="pane-tools">
            <DiagramSettingsDialog
              open={settingsHistoryKey === editor.historyKey && !editor.pending}
              theme={editor.draft.theme}
              settings={editor.draft.settings}
              diagramKind={editor.diagramKind}
              disabled={busy || !!editor.pending}
              onOpen={openSettings}
              onCancel={closeSettings}
              onApply={(theme, settings) => {
                if (busy || editor.pending || settingsHistoryKey !== editor.historyKey) return
                editor.applySettings(theme, settings)
                closeSettings()
              }}
              trigger={<Button
                ref={settingsButton}
                appearance="subtle"
                size="small"
                aria-label="Diagram settings"
                title="Diagram settings"
                disabled={busy || !!editor.pending}
                icon={<svg width="20" height="20" viewBox="0 0 24 24" fill="none" aria-hidden="true" focusable="false">
                  <path d="m9.5 3-.6 2.3-1.8 1.1-2.3-.6-2.5 4.4L4 11.8v2.1l-1.7 1.6 2.5 4.3 2.3-.6 1.8 1.1.6 2.2h5l.6-2.2 1.8-1.1 2.3.6 2.5-4.3-1.7-1.6v-2.1l1.7-1.6-2.5-4.4-2.3.6-1.8-1.1L14.5 3h-5Z"
                    stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" transform="translate(0 -1)" />
                  <circle cx="12" cy="12" r="3" stroke="currentColor" strokeWidth="1.5" />
                </svg>}
              />}
            />
            <SyntaxHelpLink />
          </div>
        </header>

        {status && <Caption1 role="status">{status}</Caption1>}
        {editor.settingsActive && <Caption1 role="status">Settings window is open.</Caption1>}
        {editor.ready ? (
          <MermaidEditor
            value={editor.draft.source}
            onChange={editor.changeSource}
            diagnostic={editor.diagnostic}
            historyKey={String(editor.historyKey)}
          />
        ) : <Spinner label="Loading selected diagram" />}

        {settingsError && (
          <MessageBar intent="error"><MessageBarBody>{settingsError}</MessageBarBody></MessageBar>
        )}
        {editor.wordError && (
          <MessageBar intent="error">
            <MessageBarBody>{editor.wordError}</MessageBarBody>
          </MessageBar>
        )}
        {editor.diagnostic && (
          <MessageBar intent="error">
            <MessageBarBody>{editor.diagnostic.message}</MessageBarBody>
          </MessageBar>
        )}

        <Dialog open={!!editor.pending} onOpenChange={(_, data) => {
          if (!data.open) editor.keepEditing()
        }}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>Discard pending edits?</DialogTitle>
              <DialogContent>
                Edits not yet written to Word will be discarded. Changes already saved in the document remain.
              </DialogContent>
              <DialogActions>
                <Button onClick={editor.keepEditing}>Keep editing</Button>
                <Button appearance="primary" disabled={!editor.pending || editor.writing}
                  onClick={editor.discardAndSwitch}>Discard and switch</Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </main>
    </FluentProvider>
  )
}
