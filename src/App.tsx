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
import { AddSquareRegular } from '@fluentui/react-icons'
import { useEffect, useState } from 'react'
import './App.css'
import { MermaidEditor } from './components/MermaidEditor'
import { renderMermaid } from './mermaid/render'
import { insertDiagram, type DiagramFormat } from './word/insertDiagram'
import { watchSelectedDiagram } from './word/selection'

const initialDiagram = `flowchart TD
    Idea[Mermaid source] --> Render[Render as SVG]
    Render --> Word[Insert into Word]
    Word --> Edit[Edit later]`

function App() {
  const [source, setSource] = useState(initialDiagram)
  const [svg, setSvg] = useState('')
  const [renderError, setRenderError] = useState('')
  const [isRendering, setIsRendering] = useState(true)
  const [isInserting, setIsInserting] = useState(false)
  const [notice, setNotice] = useState('')

  useEffect(() => {
    let stopWatching: () => void = () => undefined
    let disposed = false

    if (typeof Office !== 'undefined') {
      void Office.onReady().then(() => {
        const stop = watchSelectedDiagram(
          (payload) => {
            setSource(payload.source)
            setNotice('Selected Mermaid diagram loaded for editing.')
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
      const format: DiagramFormat = await insertDiagram(svg, source)
      setNotice(`Diagram inserted as ${format.toUpperCase()}.`)
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'Unable to insert diagram.')
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
          <Button
            appearance="primary"
            icon={<AddSquareRegular />}
            disabled={!svg || Boolean(renderError) || isRendering || isInserting}
            onClick={handleInsert}
          >
            {isInserting ? 'Inserting...' : 'Insert diagram'}
          </Button>
        </header>

        {notice && (
          <MessageBar
            intent={
              notice.startsWith('Diagram inserted') || notice.startsWith('Selected Mermaid')
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
            <MermaidEditor value={source} onChange={setSource} />
          </div>

          <div className="panel">
            <Text weight="semibold">Preview</Text>
            <div className="preview" aria-live="polite">
              {isRendering && <Spinner label="Rendering diagram" />}
              {!isRendering && renderError && (
                <MessageBar intent="error">
                  <MessageBarBody>{renderError}</MessageBarBody>
                </MessageBar>
              )}
              {!isRendering && !renderError && (
                <div
                  className="preview-svg"
                  // The SVG is sanitized after Mermaid renders it.
                  dangerouslySetInnerHTML={{ __html: svg }}
                />
              )}
            </div>
          </div>
        </section>
      </main>
    </FluentProvider>
  )
}

export default App
