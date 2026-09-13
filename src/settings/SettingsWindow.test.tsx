import { act, cleanup, render, screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { afterEach, expect, it, vi } from 'vitest'
import { DEFAULT_DIAGRAM_SETTINGS } from '../metadata/diagramSettings'
import { SettingsWindow } from './SettingsWindow'

afterEach(() => {
  cleanup()
  vi.unstubAllGlobals()
  window.history.replaceState(null, '', '/')
})

it('receives settings over the Office bridge and sends Apply without accessing Word', async () => {
  window.history.replaceState(null, '', '/?view=settings&session=test-session')
  let handler: (event: { message: string; origin: string }) => void = () => { throw new Error('Not registered') }
  const messageParent = vi.fn()
  vi.stubGlobal('Office', {
    onReady: () => Promise.resolve(),
    context: {
      officeTheme: { isDarkTheme: true },
      ui: {
        messageParent,
        addHandlerAsync: vi.fn((_event, received, callback) => {
          handler = received
          callback({ status: 'succeeded' })
        }),
      },
    },
    EventType: { DialogParentMessageReceived: 'parent' },
    AsyncResultStatus: { Succeeded: 'succeeded' },
  })
  const user = userEvent.setup()
  const { container } = render(<SettingsWindow />)
  await waitFor(() => expect(messageParent).toHaveBeenCalledWith(
    JSON.stringify({ type: 'ready', session: 'test-session' }), { targetOrigin: window.location.origin },
  ))
  expect(container.querySelector<HTMLElement>('.fui-FluentProvider')?.style.colorScheme).toBe('dark')
  act(() => handler({ origin: window.location.origin, message: JSON.stringify({
    type: 'init', session: 'test-session', theme: 'forest',
    settings: DEFAULT_DIAGRAM_SETTINGS, diagramKind: 'sequence',
  }) }))
  expect(screen.getByRole('checkbox', { name: 'Wrap text' })).toBeInTheDocument()
  await user.click(screen.getByRole('checkbox', { name: 'Wrap text' }))
  await user.click(screen.getByRole('button', { name: 'Apply' }))
  const sent = JSON.parse(messageParent.mock.calls.at(-1)![0])
  expect(sent).toMatchObject({
    type: 'apply', session: 'test-session', theme: 'forest',
    settings: { ...DEFAULT_DIAGRAM_SETTINGS, sequenceWrap: true },
  })
  expect(screen.getByRole('button', { name: 'Apply' })).toBeDisabled()
})

it('explains direct navigation without a settings session', async () => {
  vi.stubGlobal('ResizeObserver', class { observe() {} unobserve() {} disconnect() {} })
  render(<SettingsWindow />)
  expect(await screen.findByText(/gear button/)).toBeInTheDocument()
})
