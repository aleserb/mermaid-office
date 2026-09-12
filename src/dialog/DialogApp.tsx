import {
  Button,
  FluentProvider,
  MessageBar,
  MessageBarBody,
  Spinner,
  Text,
  Title2,
  webLightTheme,
} from '@fluentui/react-components'
import { useEffect, useState } from 'react'
import { DiagramPreview } from '../components/DiagramPreview'
import { MermaidEditor } from '../components/MermaidEditor'
import { SplitWorkspace } from '../components/SplitWorkspace'
import { SyntaxHelpLink } from '../components/SyntaxHelpLink'
import { ThemePicker } from '../components/ThemePicker'
import { normalizeMermaidError, type MermaidDiagnostic } from '../mermaid/diagnostics'
import { renderMermaid } from '../mermaid/render'
import type { DiagramSize, DiagramTheme } from '../metadata/payload'
import { setPreferredTheme } from '../preferences/diagramPreferences'
import { parseParentMessage, type DialogToParentMessage } from './messages'
import './dialog.css'

function sendToParent(message: DialogToParentMessage) {
  Office.context.ui.messageParent(JSON.stringify(message))
}

export function DialogApp() {
  const [source, setSource] = useState('')
  const [theme, setTheme] = useState<DiagramTheme>('default')
  const [size, setSize] = useState<DiagramSize>('medium')
  const [svg, setSvg] = useState('')
  const [error, setError] = useState<MermaidDiagnostic | null>(null)
  const [initialized, setInitialized] = useState(false)
  const [rendering, setRendering] = useState(false)

  useEffect(() => {
    Office.context.ui.addHandlerAsync(
      Office.EventType.DialogParentMessageReceived,
      (args: Office.DialogParentMessageReceivedEventArgs) => {
        try {
          const message = parseParentMessage(args.message)
          setSource(message.source)
          setTheme(message.theme)
          setSize(message.size)
          setInitialized(true)
        } catch (messageError) {
          setError(normalizeMermaidError(messageError))
        }
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          sendToParent({ type: 'ready' })
        } else {
          setError(normalizeMermaidError(result.error.message))
        }
      },
    )
  }, [])

  useEffect(() => {
    if (!initialized) {
      return
    }

    let active = true
    const timer = window.setTimeout(async () => {
      setRendering(true)
      try {
        const rendered = await renderMermaid(source, theme)
        if (active) {
          setSvg(rendered)
          setError(null)
        }
      } catch (renderError) {
        if (active) {
          setError(normalizeMermaidError(renderError, source))
        }
      } finally {
        if (active) {
          setRendering(false)
        }
      }
    }, 250)

    return () => {
      active = false
      window.clearTimeout(timer)
    }
  }, [initialized, source, theme])

  return (
    <FluentProvider theme={webLightTheme}>
      <main className="dialog-shell">
        <header className="dialog-header">
          <Title2 as="h1">Mermaid Diagram</Title2>
          <div className="dialog-actions">
            <Button onClick={() => sendToParent({ type: 'cancel' })}>Discard Changes</Button>
            <Button
              appearance="primary"
              disabled={!initialized || rendering || Boolean(error)}
              onClick={() => sendToParent({ type: 'save', source, theme, size })}
            >
              Save and Close
            </Button>
          </div>
        </header>

        {!initialized ? (
          <Spinner label="Loading diagram" />
        ) : (
          <SplitWorkspace
            className="dialog-workspace"
            ariaLabel="Mermaid diagram workspace"
            left={
              <div className="dialog-panel">
              <div className="editor-heading">
                <Text weight="semibold">Diagram source</Text>
                  <div className="diagram-options">
                    <ThemePicker
                      value={theme}
                      onChange={(value) => {
                        setTheme(value)
                        setPreferredTheme(value)
                      }}
                    />
                    <SyntaxHelpLink />
                  </div>
              </div>
                <MermaidEditor
                  value={source}
                  onChange={setSource}
                  diagnostic={error}
                />
              {error && (
                <MessageBar intent="error">
                    <MessageBarBody>{error.message}</MessageBarBody>
                </MessageBar>
              )}
              </div>
            }
            right={
              <div className="dialog-panel">
              <Text weight="semibold">Preview</Text>
              <DiagramPreview key={svg} svg={svg} loading={rendering} />
              </div>
            }
          />
        )}
      </main>
    </FluentProvider>
  )
}
