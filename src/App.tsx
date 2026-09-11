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
import { ThemePicker } from './components/ThemePicker'
import { DEFAULT_DIAGRAM } from './defaultDiagram'
import { renderMermaid } from './mermaid/render'
import { openEditorDialog } from './dialog/openEditorDialog'
import type { DiagramPayload, DiagramTheme } from './metadata/payload'
import { insertDiagram, type DiagramFormat, updateDiagram } from './word/insertDiagram'
import { watchSelectedDiagram } from './word/selection'

function App() {
  const [source, setSource] = useState(DEFAULT_DIAGRAM)
  const [theme, setTheme] = useState<DiagramTheme>('default')
  const [svg, setSvg] = useState('')
  const [renderError, setRenderError] = useState('')
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
          setRenderError('')
        }
      } catch (error) {
        if (active) {
          setRenderError(error instanceof Error ? error.message : 'Unable to render diagram.')
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
        ? await updateDiagram(svg, selectedDiagram, source, theme)
        : await insertDiagram(svg, source, theme)
      if (selectedDiagram) {
        setSelectedDiagram({ ...selectedDiagram, source, format })
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
      const result = await openEditorDialog(source, theme)
      if (result === null) {
        return
      }

      setIsInserting(true)
      const rendered = await renderMermaid(result.source, result.theme)
      const format = selectedDiagram
        ? await updateDiagram(rendered, selectedDiagram, result.source, result.theme)
        : await insertDiagram(rendered, result.source, result.theme)
      setSource(result.source)
      setTheme(result.theme)
      if (selectedDiagram) {
        setSelectedDiagram({
          ...selectedDiagram,
          source: result.source,
          theme: result.theme,
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
                  setTheme('default')
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

        <section className="workspace" aria-label="Mermaid diagram workspace">
          <div className="panel">
            <div className="editor-heading">
              <Text weight="semibold">Diagram source</Text>
              <ThemePicker value={theme} onChange={setTheme} />
            </div>
            <MermaidEditor
              value={source}
              onChange={setSource}
              diagnostic={renderError}
            />
          </div>

          <div className="panel">
            <Text weight="semibold">Preview</Text>
            <DiagramPreview svg={svg} loading={isRendering} error={renderError} />
          </div>
        </section>
      </main>
    </FluentProvider>
  )
}

export default App
