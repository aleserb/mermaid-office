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
import { ThemePicker } from '../components/ThemePicker'
import { usePaneEditor } from './usePaneEditor'
import './pane.css'

export function PaneApp() {
  const editor = usePaneEditor()
  const busy = !editor.ready || editor.writing || editor.loadingSelection
  const status = editor.writing
    ? 'Writing diagram to Word...'
    : editor.target
      ? editor.dirty ? 'Changes not yet written to Word.' : 'Diagram is up to date in Word.'
      : 'Place the cursor in Word and insert your diagram to start live updates.'

  return (
    <FluentProvider theme={webLightTheme} style={{ colorScheme: 'light' }}>
      <main className="pane-shell">
        <header className="pane-header">
          <Caption1 className="pane-build">Build {__BUILD_VERSION__} - Experimental pane</Caption1>
          <div className="diagram-options">
            <SyntaxHelpLink />
            <ThemePicker value={editor.draft.theme} onChange={editor.changeTheme} />
          </div>
          <div className="pane-actions">
            <Button disabled={busy} onClick={editor.newDiagram}>New diagram</Button>
            <Button disabled={busy} onClick={() => void editor.loadSelected()}>Edit selected</Button>
            {!editor.target && (
              <Button appearance="primary" disabled={!editor.canInsert} onClick={() => void editor.insert()}>
                Insert diagram
              </Button>
            )}
            {editor.canRetry && (
              <Button disabled={busy} onClick={editor.retry}>Retry update</Button>
            )}
          </div>
        </header>

        <Caption1 role="status">{status}</Caption1>
        {editor.target && (
          <Caption1>Live updates stay linked to this diagram. Use Edit selected to switch diagrams.</Caption1>
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
        {editor.ready ? (
          <MermaidEditor
            value={editor.draft.source}
            onChange={editor.changeSource}
            diagnostic={editor.diagnostic}
            historyKey={String(editor.historyKey)}
          />
        ) : <Spinner label="Loading selected diagram" />}

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
