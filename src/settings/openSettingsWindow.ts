import type { DiagramSettings } from '../metadata/diagramSettings'
import type { DiagramTheme } from '../metadata/payload'
import { parseSettingsMessage, type SettingsSnapshot } from './messages'

interface SettingsWindowCallbacks {
  onApply: (theme: DiagramTheme, settings: DiagramSettings) => void
  onClose: () => void
  onError: (message: string) => void
}

export function openSettingsWindow(snapshot: SettingsSnapshot, callbacks: SettingsWindowCallbacks): { dispose: () => void } {
  const session = crypto.randomUUID()
  const url = new URL('./', window.location.href)
  const version = new URLSearchParams(window.location.search).get('v')
  if (version) url.searchParams.set('v', version)
  url.searchParams.set('view', 'settings')
  url.searchParams.set('session', session)
  const initial = JSON.stringify({ ...snapshot, type: 'init', session })
  let dialog: Office.Dialog | undefined
  let finished = false
  let initialized = false
  let timeout: number | undefined

  const finish = (error?: string) => {
    if (finished) return
    finished = true
    window.clearTimeout(timeout)
    window.removeEventListener('pagehide', dispose)
    try {
      dialog?.close()
      if (error) callbacks.onError(error)
    } catch (closeError) {
      callbacks.onError(closeError instanceof Error ? closeError.message : 'Unable to close the settings window.')
    } finally {
      callbacks.onClose()
    }
  }
  const dispose = () => {
    if (finished) return
    finished = true
    window.clearTimeout(timeout)
    window.removeEventListener('pagehide', dispose)
    try {
      dialog?.close()
    } catch (error) {
      console.error('Unable to close the settings window during pane cleanup.', error)
    }
  }

  if (!Office.context.requirements?.isSetSupported('DialogApi', '1.2')) {
    finish('This version of Word cannot open the settings window. Update Word and try again.')
    return { dispose }
  }

  window.addEventListener('pagehide', dispose)
  timeout = window.setTimeout(() => finish('The settings window did not load. Close and reopen settings to try again.'), 30000)
  try {
    Office.context.ui.displayDialogAsync(url.href, {
      displayInIframe: true,
      width: Math.min(90, Math.max(30, Math.round(560 / (window.screen.width || window.innerWidth) * 100))),
      height: Math.min(90, Math.max(45, Math.round(740 / (window.screen.height || window.innerHeight) * 100))),
    }, result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        finish(`Unable to open settings: ${result.error.message} (${result.error.code}).`)
        return
      }
      dialog = result.value
      if (finished) {
        dialog.close()
        return
      }
      try {
        dialog.addEventHandler(Office.EventType.DialogEventReceived, event => {
          if ('error' in event) {
            dialog = undefined
            finish(event.error === 12006 ? undefined : `The settings window closed with an Office error (${event.error}).`)
          }
        })
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, event => {
          if (finished || !('message' in event)) return
          if (event.origin && event.origin !== url.origin) {
            console.warn('Ignored a settings message from an unexpected origin.')
            return
          }
          try {
            const message = parseSettingsMessage(event.message)
            if (message.session !== session) {
              console.warn('Ignored a message from another settings session.')
              return
            }
            if (message.type === 'ready') {
              dialog?.messageChild(initial, { targetOrigin: url.origin })
              initialized = true
              window.clearTimeout(timeout)
            } else if (message.type === 'cancel') {
              finish()
            } else if (message.type === 'apply' && initialized) {
              callbacks.onApply(message.theme, message.settings)
              finish()
            } else {
              finish('The settings window sent an unexpected message. Please reopen it.')
            }
          } catch (error) {
            finish(error instanceof Error ? error.message : 'Unable to read settings from the window.')
          }
        })
      } catch (error) {
        finish(error instanceof Error ? error.message : 'Unable to connect to the settings window.')
      }
    })
  } catch (error) {
    finish(error instanceof Error ? error.message : 'Unable to open the settings window.')
  }
  return { dispose }
}
