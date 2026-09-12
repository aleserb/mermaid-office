import { afterEach, expect, it, vi } from 'vitest'
import { openEditorDialog } from './openEditorDialog'
import { parseParentMessage } from './messages'

afterEach(() => vi.unstubAllGlobals())

it.each([true, false])('forwards Word dark mode %s without changing the Mermaid theme', async (darkMode) => {
  const handlers = new Map<string, (args: { message: string }) => void>()
  const dialog = {
    close: vi.fn(),
    messageChild: vi.fn(),
    addEventHandler: (event: string, handler: (args: { message: string }) => void) => {
      handlers.set(event, handler)
    },
  }
  vi.stubGlobal('Office', {
    context: {
      officeTheme: { isDarkTheme: darkMode },
      ui: {
        displayDialogAsync: (
          _url: string,
          _options: object,
          callback: (result: { status: string; value: typeof dialog }) => void,
        ) => callback({ status: 'succeeded', value: dialog }),
      },
    },
    AsyncResultStatus: { Failed: 'failed' },
    EventType: { DialogMessageReceived: 'message', DialogEventReceived: 'event' },
  })

  const result = openEditorDialog('sequenceDiagram', 'forest', 'medium', 'update')
  const receive = handlers.get('message')
  if (!receive) throw new Error('Dialog message handler was not registered.')
  receive({ message: JSON.stringify({ type: 'ready' }) })
  expect(dialog.messageChild).toHaveBeenCalledOnce()
  expect(parseParentMessage(dialog.messageChild.mock.calls[0][0])).toEqual({
    type: 'initialize',
    source: 'sequenceDiagram',
    theme: 'forest',
    size: 'medium',
    mode: 'update',
    darkMode,
  })
  receive({ message: JSON.stringify({ type: 'cancel' }) })
  expect(await result).toBeNull()
  expect(dialog.close).toHaveBeenCalledOnce()
})
