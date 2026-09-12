import {
  Button,
  FluentProvider,
  MessageBar,
  MessageBarBody,
  Text,
  webLightTheme,
} from '@fluentui/react-components'
import { AddSquareRegular } from '@fluentui/react-icons'
import { OpenRegular } from '@fluentui/react-icons'
import { useEffect, useState } from 'react'
import './App.css'
import { DiagramPreview } from './components/DiagramPreview'
import { MermaidEditor } from './components/MermaidEditor'
import { SplitWorkspace } from './components/SplitWorkspace'
import { SyntaxHelpLink } from './components/SyntaxHelpLink'
import { ThemePicker } from './components/ThemePicker'
import { DEFAULT_DIAGRAM } from './defaultDiagram'
import { normalizeMermaidError, type MermaidDiagnostic } from './mermaid/diagnostics'
import { renderMermaid } from './mermaid/render'
import { openEditorDialog } from './dialog/openEditorDialog'
import type { DiagramPayload, DiagramSize, DiagramTheme } from './metadata/payload'
import {
  getPreferredTheme,
  setPreferredTheme,
} from './preferences/diagramPreferences'
import { insertDiagram, type DiagramFormat, updateDiagram } from './word/insertDiagram'
import { watchSelectedDiagram } from './word/selection'

function App() {
  const [source, setSource] = useState(DEFAULT_DIAGRAM)
  const [theme, setTheme] = useState<DiagramTheme>(getPreferredTheme)
  const [size, setSize] = useState<DiagramSize>('medium')
  const [svg, setSvg] = useState('')
  const [renderError, setRenderError] = useState<MermaidDiagnostic | null>(null)
  const [isRendering, setIsRendering] = useState(true)
  const [isInserting, setIsInserting] = useState(false)
  const [notice, setNotice] = useState('')
  const [selectedDiagram, setSelectedDiagram] = useState<DiagramPayload | null>(null)

  useEffect(() => {
    let stopWatching: () => void = () => undefined
    let disposed = false

    if (typeof Office !== 'undefined') {
      void Office.onReady().then(() => {
        const stop = watchSelectedDiagram(
          (payload) => {
            setSelectedDiagram(payload)
            if (payload) {
              setSource(payload.source)
              setTheme(payload.theme)
              setSize(payload.size)
              setNotice('Selected Mermaid diagram loaded for editing.')
            }
          },
          (error) => {
            setNotice(error.message)
          },
        )
        if (disposed) {
          stop()
        } else {
          stopWatching = stop
        }
      })
    }

    return () => {
      disposed = true
      stopWatching()
    }
  }, [])

  useEffect(() => {
    let active = true
    const timer = window.setTimeout(async () => {
      setIsRendering(true)
      try {
        const rendered = await renderMermaid(source, theme)
        if (active) {
          setSvg(rendered)
          setRenderError(null)
        }
      } catch (error) {
        if (active) {
          setRenderError(normalizeMermaidError(error, source))
        }
      } finally {
        if (active) {
          setIsRendering(false)
        }
      }
    }, 250)

    return () => {
      active = false
      window.clearTimeout(timer)
    }
  }, [source, theme])

  const handleInsert = async () => {
    if (!svg || renderError) {
      return
    }

    setIsInserting(true)
    setNotice('')
    try {
      const format: DiagramFormat = selectedDiagram
        ? await updateDiagram(
            svg,
            selectedDiagram,
            source,
            theme,
            size,
            size !== selectedDiagram.size,
          )
        : await insertDiagram(svg, source, theme, size)
      if (selectedDiagram) {
        setSelectedDiagram({ ...selectedDiagram, source, theme, size, format })
      }
      setNotice(
        selectedDiagram
          ? `Diagram updated as ${format.toUpperCase()}.`
          : `Diagram inserted as ${format.toUpperCase()}.`,
      )
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'Unable to insert diagram.')
    } finally {
      setIsInserting(false)
    }
  }

  const handleOpenDialog = async () => {
    setNotice('')
    try {
      const result = await openEditorDialog(source, theme, size)
      if (result === null) {
        return
      }

      setIsInserting(true)
      const rendered = await renderMermaid(result.source, result.theme)
      const format = selectedDiagram
        ? await updateDiagram(
            rendered,
            selectedDiagram,
            result.source,
            result.theme,
            result.size,
            result.size !== selectedDiagram.size,
          )
        : await insertDiagram(rendered, result.source, result.theme, result.size)
      setSource(result.source)
      setTheme(result.theme)
      setSize(result.size)
      setPreferredTheme(result.theme)
      if (selectedDiagram) {
        setSelectedDiagram({
          ...selectedDiagram,
          source: result.source,
          theme: result.theme,
          size: result.size,
          format,
        })
      }
      setNotice(
        selectedDiagram
          ? `Diagram updated as ${format.toUpperCase()}.`
          : `Diagram inserted as ${format.toUpperCase()}.`,
      )
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'Unable to open the editor.')
    } finally {
      setIsInserting(false)
    }
  }

  return (
    <FluentProvider theme={webLightTheme}>
      <main className="app-shell">
        <header className="app-header">
          <div className="header-actions">
            {selectedDiagram && (
              <Button
                disabled={isInserting}
                onClick={() => {
                  setSelectedDiagram(null)
                  setSource(DEFAULT_DIAGRAM)
                  setTheme(getPreferredTheme())
                  setSize('medium')
                  setNotice('Ready to insert a new diagram.')
                }}
              >
                New diagram
              </Button>
            )}
            <Button
              icon={<OpenRegular />}
              disabled={isInserting}
              onClick={handleOpenDialog}
            >
              Expand editor
            </Button>
            <Button
              appearance="primary"
              icon={<AddSquareRegular />}
              disabled={!svg || Boolean(renderError) || isRendering || isInserting}
              onClick={handleInsert}
            >
              {isInserting
                ? selectedDiagram
                  ? 'Updating...'
                  : 'Inserting...'
                : selectedDiagram
                  ? 'Update diagram'
                  : 'Insert diagram'}
            </Button>
          </div>
        </header>

        {notice && (
          <MessageBar
            intent={
              notice.startsWith('Diagram inserted') ||
              notice.startsWith('Diagram updated') ||
              notice.startsWith('Selected Mermaid')
                ? 'success'
                : 'error'
            }
          >
            <MessageBarBody>{notice}</MessageBarBody>
          </MessageBar>
        )}

        <SplitWorkspace
          className="workspace"
          ariaLabel="Mermaid diagram workspace"
          left={
            <div className="panel">
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
                diagnostic={renderError}
                historyKey={selectedDiagram?.id ?? 'new-diagram'}
              />
            </div>
          }
          right={
            <div className="panel">
              <Text weight="semibold">Preview</Text>
              <DiagramPreview
                key={svg}
                svg={svg}
                loading={isRendering}
                error={renderError?.message}
              />
            </div>
          }
        />
      </main>
    </FluentProvider>
  )
}

export default App
