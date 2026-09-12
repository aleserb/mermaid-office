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
import { useState } from 'react'
import { MermaidEditor } from '../components/MermaidEditor'
import { EditorDisplayControls } from '../components/EditorDisplayControls'
import { DEFAULT_EDITOR_DISPLAY } from '../components/editorDisplay'
import { SyntaxHelpLink } from '../components/SyntaxHelpLink'
import { ThemePicker } from '../components/ThemePicker'
import { usePaneEditor } from './usePaneEditor'
import './pane.css'

export function PaneApp() {
  const editor = usePaneEditor()
  const [display, setDisplay] = useState(DEFAULT_EDITOR_DISPLAY)
  const busy = !editor.ready || editor.writing || editor.loadingSelection
  const status = editor.writing
    ? 'Writing diagram to Word...'
    : editor.target && editor.dirty
      ? 'Changes not yet written to Word.'
      : ''

  return (
    <FluentProvider theme={webLightTheme} style={{ colorScheme: 'light' }}>
      <main className="pane-shell">
        <header className="pane-header">
          <div className="diagram-options">
            <ThemePicker value={editor.draft.theme} onChange={editor.changeTheme} />
          </div>
          <EditorDisplayControls value={display} onChange={setDisplay} />
        </header>

        {status && <Caption1 role="status">{status}</Caption1>}
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
        {editor.ready ? (
          <MermaidEditor
            value={editor.draft.source}
            onChange={editor.changeSource}
            diagnostic={editor.diagnostic}
            historyKey={String(editor.historyKey)}
            display={display}
          />
        ) : <Spinner label="Loading selected diagram" />}

        <footer className="pane-footer">
          {(!editor.target || editor.canRetry) && <div className="pane-actions">
            {!editor.target && (
              <Button appearance="primary" disabled={!editor.canInsert} onClick={() => void editor.insert()}>
                Insert diagram
              </Button>
            )}
            {editor.canRetry && (
              <Button disabled={busy} onClick={editor.retry}>Retry update</Button>
            )}
          </div>}
          <SyntaxHelpLink />
        </footer>

        {editor.pending && <Dialog open onOpenChange={(_, data) => {
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
                <Button appearance="primary" onClick={editor.discardAndSwitch}>Discard and switch</Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>}
      </main>
    </FluentProvider>
  )
}
