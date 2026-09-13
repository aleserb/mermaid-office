import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { DEFAULT_DIAGRAM_SETTINGS } from '../metadata/diagramSettings'
import { openSettingsWindow } from './openSettingsWindow'

const snapshot = { theme: 'forest' as const, settings: DEFAULT_DIAGRAM_SETTINGS, diagramKind: 'flowchart' as const }
const callbacks = { onApply: vi.fn(), onClose: vi.fn(), onError: vi.fn() }
let handlers: Record<string, (event: { message: string; origin?: string } | { error: number }) => void>
const dialog = { close: vi.fn(), messageChild: vi.fn(), addEventHandler: vi.fn() }
const display = vi.fn()
const supported = vi.fn()

beforeEach(() => {
  vi.useFakeTimers()
  vi.spyOn(window.screen, 'width', 'get').mockReturnValue(1920)
  vi.spyOn(window.screen, 'height', 'get').mockReturnValue(1080)
  handlers = {}
  vi.resetAllMocks()
  supported.mockReturnValue(true)
  dialog.addEventHandler.mockImplementation((type, handler) => { handlers[type] = handler })
  vi.stubGlobal('Office', {
    context: { ui: { displayDialogAsync: display }, requirements: { isSetSupported: supported } },
    AsyncResultStatus: { Succeeded: 'succeeded' },
    EventType: { DialogMessageReceived: 'message', DialogEventReceived: 'event' },
  })
})

afterEach(() => { vi.useRealTimers(); vi.unstubAllGlobals(); vi.restoreAllMocks() })

function opened() {
  display.mock.calls[0][2]({ status: 'succeeded', value: dialog })
  return new URL(display.mock.calls[0][0]).searchParams.get('session')
}

function message(session: string | null, type: string, extra = {}) {
  handlers.message({ message: JSON.stringify({ session, type, ...extra }), origin: window.location.origin })
}

it('opens over Word and exchanges settings only after a ready handshake', () => {
  openSettingsWindow(snapshot, callbacks)
  const [url, options] = display.mock.calls[0]
  expect(new URL(url).searchParams.get('view')).toBe('settings')
  expect(url).not.toContain('forest')
  expect(options.displayInIframe).toBe(true)
  expect(options).toMatchObject({ width: 30, height: 76 })
  const session = opened()
  expect(dialog.messageChild).not.toHaveBeenCalled()
  message(session, 'ready')
  expect(dialog.messageChild).toHaveBeenCalledWith(
    JSON.stringify({ ...snapshot, type: 'init', session }), { targetOrigin: window.location.origin },
  )
  const settings = { ...DEFAULT_DIAGRAM_SETTINGS, imageQuality: 'high' }
  message(session, 'apply', { theme: 'dark', settings })
  message(session, 'apply', { theme: 'dark', settings })
  expect(callbacks.onApply).toHaveBeenCalledExactlyOnceWith('dark', settings)
  expect(callbacks.onClose).toHaveBeenCalledOnce()
  expect(dialog.close).toHaveBeenCalledOnce()
  expect(callbacks.onError).not.toHaveBeenCalled()
})

it.each(['cancel', 'title-bar'])('closes without applying on %s', kind => {
  openSettingsWindow(snapshot, callbacks)
  const session = opened()
  if (kind === 'cancel') message(session, 'cancel')
  else handlers.event({ error: 12006 })
  expect(callbacks.onClose).toHaveBeenCalledOnce()
  expect(callbacks.onApply).not.toHaveBeenCalled()
  expect(callbacks.onError).not.toHaveBeenCalled()
})

it('ignores other origins and sessions and rejects invalid settings', () => {
  vi.spyOn(console, 'warn').mockImplementation(() => {})
  openSettingsWindow(snapshot, callbacks)
  const session = opened()
  handlers.message({ message: JSON.stringify({ session, type: 'ready' }), origin: 'https://untrusted.example' })
  message('different-session', 'ready')
  expect(dialog.messageChild).not.toHaveBeenCalled()
  message(session, 'ready')
  message(session, 'apply', { theme: 'forest', settings: { imageQuality: 'unlimited' } })
  expect(callbacks.onApply).not.toHaveBeenCalled()
  expect(callbacks.onError).toHaveBeenCalledWith('Invalid diagram settings.')
  expect(callbacks.onClose).toHaveBeenCalledOnce()
  vi.restoreAllMocks()
})

it('closes a late popup after the pane was disposed without applying or invoking callbacks', () => {
  const handle = openSettingsWindow(snapshot, callbacks)
  handle.dispose()
  opened()
  expect(dialog.close).toHaveBeenCalledOnce()
  expect(callbacks.onClose).not.toHaveBeenCalled()
  expect(dialog.addEventHandler).not.toHaveBeenCalled()
})

it('reports blocked popups and unsupported Office versions without freezing the pane', () => {
  openSettingsWindow(snapshot, callbacks)
  display.mock.calls[0][2]({ status: 'failed', error: { code: 12009, message: 'Popup blocked' } })
  expect(callbacks.onError).toHaveBeenCalledWith(expect.stringContaining('Popup blocked'))
  expect(callbacks.onClose).toHaveBeenCalledOnce()
  supported.mockReturnValue(false)
  openSettingsWindow(snapshot, callbacks)
  expect(callbacks.onError).toHaveBeenLastCalledWith(expect.stringContaining('Update Word'))
  expect(display).toHaveBeenCalledOnce()
})

it('times out an unresponsive child and closes it', () => {
  openSettingsWindow(snapshot, callbacks)
  opened()
  vi.advanceTimersByTime(30000)
  expect(callbacks.onError).toHaveBeenCalledWith(expect.stringContaining('did not load'))
  expect(dialog.close).toHaveBeenCalledOnce()
  expect(callbacks.onClose).toHaveBeenCalledOnce()
})

it('surfaces event-registration errors and closes when the pane unloads', () => {
  dialog.addEventHandler.mockImplementationOnce(() => { throw new Error('Registration failed') })
  openSettingsWindow(snapshot, callbacks)
  opened()
  expect(callbacks.onError).toHaveBeenCalledWith('Registration failed')
  expect(callbacks.onClose).toHaveBeenCalledOnce()
  expect(dialog.close).toHaveBeenCalledOnce()
  vi.clearAllMocks()
  openSettingsWindow(snapshot, callbacks)
  opened()
  window.dispatchEvent(new Event('pagehide'))
  expect(dialog.close).toHaveBeenCalledOnce()
  expect(callbacks.onClose).not.toHaveBeenCalled()
})
