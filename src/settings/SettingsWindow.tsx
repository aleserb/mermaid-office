import { useEffect, useRef, useState } from 'react'
import { FluentProvider, MessageBar, MessageBarBody, Spinner, webLightTheme } from '@fluentui/react-components'
import { DiagramSettingsDialog } from '../components/DiagramSettingsDialog'
import { parseSettingsMessage, type SettingsMessage, type SettingsSnapshot } from './messages'
import './settingsWindow.css'

export function SettingsWindow() {
  const [snapshot, setSnapshot] = useState<SettingsSnapshot | null>(null)
  const [error, setError] = useState('')
  const [submitted, setSubmitted] = useState(false)
  const initialized = useRef(false)
  const session = new URLSearchParams(window.location.search).get('session') ?? ''
  const send = (message: SettingsMessage) => {
    try {
      Office.context.ui.messageParent(JSON.stringify(message), { targetOrigin: window.location.origin })
    } catch (error) {
      setError(error instanceof Error ? error.message : 'Unable to contact the Mermaid pane.')
      setSubmitted(false)
    }
  }

  useEffect(() => {
    document.title = 'Mermaid diagram settings'
    let active = true
    let interval: number | undefined
    let timeout: number | undefined
    const handler = (event: Office.DialogParentMessageReceivedEventArgs) => {
      if (!active || initialized.current) return
      if (event.origin && event.origin !== window.location.origin) {
        console.warn('Ignored settings initialization from an unexpected origin.')
        return
      }
      try {
        const message = parseSettingsMessage(event.message)
        if (message.type !== 'init' || message.session !== session) throw new Error('Invalid settings window initialization.')
        initialized.current = true
        window.clearInterval(interval)
        window.clearTimeout(timeout)
        setSnapshot(message)
        setError('')
      } catch (error) {
        setError(error instanceof Error ? error.message : 'Unable to initialize settings.')
      }
    }
    void (async () => {
      try {
        if (!session || typeof Office === 'undefined') throw new Error('Open settings using the gear button in the Mermaid pane.')
        await Office.onReady()
        if (!active) return
        Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, handler, result => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            if (active) setError(`Unable to receive settings: ${result.error.message}`)
            return
          }
          if (!active) return
          const ready = () => send({ type: 'ready', session })
          // Retry the handshake until the pane has attached its message handler.
          interval = window.setInterval(ready, 500)
          timeout = window.setTimeout(() => {
            window.clearInterval(interval)
            setError('The Mermaid pane did not respond. Close this window and reopen settings.')
          }, 30000)
          ready()
        })
      } catch (error) {
        if (active) setError(error instanceof Error ? error.message : 'Unable to initialize the settings window.')
      }
    })()
    return () => {
      active = false
      window.clearInterval(interval)
      window.clearTimeout(timeout)
      // Office.UI has no remove-handler API; the active guard retires this listener.
    }
  }, [session])

  return (
    <FluentProvider theme={webLightTheme} className="settings-window">
      {error ? <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar> : snapshot ? (
        <DiagramSettingsDialog
          theme={snapshot.theme}
          settings={snapshot.settings}
          diagramKind={snapshot.diagramKind}
          disabled={submitted}
          onCancel={() => { setSubmitted(true); send({ type: 'cancel', session }) }}
          onApply={(theme, settings) => {
            setSubmitted(true)
            send({ type: 'apply', session, theme, settings })
          }}
        />
      ) : <Spinner label="Loading diagram settings" />}
    </FluentProvider>
  )
}
