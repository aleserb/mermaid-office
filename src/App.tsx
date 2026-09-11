import {
  Button,
  FluentProvider,
  MessageBar,
  MessageBarBody,
  Text,
  Title2,
  webLightTheme,
} from '@fluentui/react-components'
import { AddSquareRegular } from '@fluentui/react-icons'
import { OpenRegular } from '@fluentui/react-icons'
import { useEffect, useState } from 'react'
import './App.css'
import { DiagramPreview } from './components/DiagramPreview'
import { MermaidEditor } from './components/MermaidEditor'
import { DEFAULT_DIAGRAM } from './defaultDiagram'
import { renderMermaid } from './mermaid/render'
import { openEditorDialog } from './dialog/openEditorDialog'
import type { DiagramPayload } from './metadata/payload'
import { insertDiagram, type DiagramFormat, updateDiagram } from './word/insertDiagram'
import { watchSelectedDiagram } from './word/selection'

function App() {
  const [source, setSource] = useState(DEFAULT_DIAGRAM)
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
        const rendered = await renderMermaid(source)
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
  }, [source])

  const handleInsert = async () => {
    if (!svg || renderError) {
      return
    }

    setIsInserting(true)
    setNotice('')
    try {
      const format: DiagramFormat = selectedDiagram
        ? await updateDiagram(svg, selectedDiagram, source)
        : await insertDiagram(svg, source)
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
      const editedSource = await openEditorDialog(source)
      if (editedSource === null) {
        return
      }

      setIsInserting(true)
      const rendered = await renderMermaid(editedSource)
      const format = selectedDiagram
        ? await updateDiagram(rendered, selectedDiagram, editedSource)
        : await insertDiagram(rendered, editedSource)
      setSource(editedSource)
      if (selectedDiagram) {
        setSelectedDiagram({ ...selectedDiagram, source: editedSource, format })
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
          <div>
            <Title2 as="h1">Mermaid Office</Title2>
            <Text block>Write Mermaid and insert a crisp diagram into Word.</Text>
          </div>
          <div className="header-actions">
            {selectedDiagram && (
              <Button
                disabled={isInserting}
                onClick={() => {
                  setSelectedDiagram(null)
                  setSource(DEFAULT_DIAGRAM)
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
            <Text weight="semibold">Diagram source</Text>
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
