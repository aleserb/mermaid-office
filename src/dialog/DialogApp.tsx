import {
  Button,
  FluentProvider,
  MessageBar,
  MessageBarBody,
  Spinner,
  Text,
  webDarkTheme,
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
import { systemUsesDarkTheme } from '../preferences/officeTheme'
import { rasterizeSvg } from '../word/insertDiagram'
import { parseParentMessage, type DialogToParentMessage } from './messages'
import './dialog.css'

function sendToParent(message: DialogToParentMessage) {
  Office.context.ui.messageParent(JSON.stringify(message))
}

export function DialogApp() {
  const [darkMode, setDarkMode] = useState(systemUsesDarkTheme)
  const [source, setSource] = useState('')
  const [theme, setTheme] = useState<DiagramTheme>('default')
  const [size, setSize] = useState<DiagramSize>('medium')
  const [mode, setMode] = useState<'insert' | 'update'>('insert')
  const [svg, setSvg] = useState('')
  const [error, setError] = useState<MermaidDiagnostic | null>(null)
  const [initialized, setInitialized] = useState(false)
  const [rendering, setRendering] = useState(false)
  const [saving, setSaving] = useState(false)

  useEffect(() => {
    Office.context.ui.addHandlerAsync(
      Office.EventType.DialogParentMessageReceived,
      (args: Office.DialogParentMessageReceivedEventArgs) => {
        try {
          const message = parseParentMessage(args.message)
          setDarkMode(message.darkMode ?? systemUsesDarkTheme())
          setSource(message.source)
          setTheme(message.theme)
          setSize(message.size)
          setMode(message.mode)
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

  const saveDiagram = async () => {
    if (!svg || error) {
      return
    }

    setSaving(true)
    try {
      const raster = await rasterizeSvg(svg, size)
      sendToParent({ type: 'save', source, theme, size, svg, raster })
    } catch (saveError) {
      setError(normalizeMermaidError(saveError))
      setSaving(false)
    }
  }

  return (
    <FluentProvider
      theme={darkMode ? webDarkTheme : webLightTheme}
      style={{ colorScheme: darkMode ? 'dark' : 'light' }}
    >
      <main className="dialog-shell">
        <header className="dialog-header">
          <div className="diagram-options">
            <Text className="build-version" size={200}>
              Build {__BUILD_VERSION__}
            </Text>
            <SyntaxHelpLink />
            <ThemePicker
              value={theme}
              onChange={(value) => {
                setTheme(value)
                setPreferredTheme(value)
              }}
            />
          </div>
          <div className="dialog-actions">
            <Button onClick={() => sendToParent({ type: 'cancel' })}>Discard Changes</Button>
            <Button
              appearance="primary"
              disabled={!initialized || rendering || saving || Boolean(error)}
              onClick={saveDiagram}
            >
              {saving
                ? mode === 'update'
                  ? 'Updating...'
                  : 'Inserting...'
                : mode === 'update'
                  ? 'Update'
                  : 'Insert'}
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
                <div className="dialog-editor-body">
                  <MermaidEditor
                    darkMode={darkMode}
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
              </div>
            }
            right={
              <div className="dialog-panel">
                <DiagramPreview key={svg} svg={svg} loading={rendering} />
              </div>
            }
          />
        )}
      </main>
    </FluentProvider>
  )
}
