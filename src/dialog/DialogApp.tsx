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
import { renderMermaid } from '../mermaid/render'
import { parseParentMessage, type DialogToParentMessage } from './messages'
import './dialog.css'

function sendToParent(message: DialogToParentMessage) {
  Office.context.ui.messageParent(JSON.stringify(message))
}

export function DialogApp() {
  const [source, setSource] = useState('')
  const [svg, setSvg] = useState('')
  const [error, setError] = useState('')
  const [initialized, setInitialized] = useState(false)
  const [rendering, setRendering] = useState(false)

  useEffect(() => {
    Office.context.ui.addHandlerAsync(
      Office.EventType.DialogParentMessageReceived,
      (args: Office.DialogParentMessageReceivedEventArgs) => {
        try {
          const message = parseParentMessage(args.message)
          setSource(message.source)
          setInitialized(true)
        } catch (messageError) {
          setError(
            messageError instanceof Error
              ? messageError.message
              : 'Unable to initialize the editor.',
          )
        }
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          sendToParent({ type: 'ready' })
        } else {
          setError(result.error.message)
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
        const rendered = await renderMermaid(source)
        if (active) {
          setSvg(rendered)
          setError('')
        }
      } catch (renderError) {
        if (active) {
          setError(
            renderError instanceof Error ? renderError.message : 'Unable to render diagram.',
          )
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
  }, [initialized, source])

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
              onClick={() => sendToParent({ type: 'save', source })}
            >
              Save and Close
            </Button>
          </div>
        </header>

        {!initialized ? (
          <Spinner label="Loading diagram" />
        ) : (
          <section className="dialog-workspace">
            <div className="dialog-panel">
              <Text weight="semibold">Diagram source</Text>
              <MermaidEditor value={source} onChange={setSource} diagnostic={error} />
              {error && (
                <MessageBar intent="error">
                  <MessageBarBody>{error}</MessageBarBody>
                </MessageBar>
              )}
            </div>
            <div className="dialog-panel">
              <Text weight="semibold">Preview</Text>
              <DiagramPreview svg={svg} loading={rendering} />
            </div>
          </section>
        )}
      </main>
    </FluentProvider>
  )
}
