import { StrictMode } from 'react'
import { afterEach, expect, it, vi } from 'vitest'

const root = vi.hoisted(() => ({ render: vi.fn() }))
vi.mock('react-dom/client', () => ({ createRoot: vi.fn(() => root) }))

afterEach(() => {
  root.render.mockClear()
  vi.resetModules()
  document.body.replaceChildren()
  window.history.replaceState(null, '', '/')
})

it.each(['/', '/?view=pane'])('opens the code-only pane at %s', async (url) => {
  window.history.replaceState(null, '', url)
  const container = document.createElement('div')
  container.id = 'root'
  document.body.append(container)
  const { PaneApp } = await import('./pane/PaneApp')
  await import('./main')
  expect(root.render).toHaveBeenCalledExactlyOnceWith(
    <StrictMode><PaneApp /></StrictMode>,
  )
})
