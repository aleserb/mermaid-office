import { cleanup, render, screen, waitFor } from '@testing-library/react'
import { webLightTheme } from '@fluentui/react-components'
import { afterEach, expect, it, vi } from 'vitest'
import { renderMermaid } from '../mermaid/render'
import { DialogApp } from './DialogApp'

vi.mock('../mermaid/render', () => ({
  renderMermaid: vi.fn().mockResolvedValue('<svg></svg>'),
}))
vi.mock('../word/insertDiagram', () => ({
  rasterizeSvg: vi.fn(),
}))

afterEach(() => {
  cleanup()
  vi.unstubAllGlobals()
  vi.clearAllMocks()
})

it('keeps the dialog light even with dark host, system, and legacy message preferences', async () => {
  vi.stubGlobal('__BUILD_VERSION__', 'test')
  vi.stubGlobal('matchMedia', vi.fn().mockReturnValue({
    matches: true,
    addEventListener: vi.fn(),
    removeEventListener: vi.fn(),
  }))
  vi.stubGlobal('Office', {
    context: {
      officeTheme: { isDarkTheme: true },
      ui: {
        addHandlerAsync: (
          _event: string,
          handler: (args: { message: string }) => void,
          callback: (result: { status: string }) => void,
        ) => {
          handler({
            message: JSON.stringify({
              type: 'initialize',
              mode: 'insert',
              source: 'flowchart LR\nA-->B',
              theme: 'dark',
              size: 'medium',
              darkMode: true,
            }),
          })
          callback({ status: 'succeeded' })
        },
        messageParent: vi.fn(),
      },
    },
    AsyncResultStatus: { Succeeded: 'succeeded' },
    EventType: { DialogParentMessageReceived: 'parent-message' },
  })

  const { container } = render(<DialogApp />)
  const provider = container.querySelector<HTMLElement>('.fui-FluentProvider')
  if (!provider) throw new Error('Dialog theme provider was not rendered.')
  expect(provider.style.colorScheme).toBe('light')
  expect(getComputedStyle(provider).getPropertyValue('--colorNeutralBackground1'))
    .toBe(webLightTheme.colorNeutralBackground1)
  expect(screen.getByRole('combobox', { name: 'Diagram theme' })).toHaveValue('dark')
  await waitFor(() => expect(renderMermaid).toHaveBeenCalledWith('flowchart LR\nA-->B', 'dark'))
})
