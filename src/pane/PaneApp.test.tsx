import { act, cleanup, render, screen, within } from '@testing-library/react'
import { afterEach, expect, it, vi } from 'vitest'
import { PaneApp } from './PaneApp'
import { insertDiagramWithPayload, updateDiagramById } from '../word/insertDiagram'
import { createDiagramPayload } from '../metadata/payload'
import { watchSelectedDiagram } from '../word/selection'

vi.mock('../mermaid/render', () => ({
  renderMermaid: vi.fn().mockResolvedValue('<svg></svg>'),
}))
vi.mock('../word/insertDiagram', () => ({
  insertDiagramWithPayload: vi.fn(),
  updateDiagramById: vi.fn(),
}))
vi.mock('../word/selection', () => ({
  getSelectedDiagram: vi.fn().mockResolvedValue(null),
  watchSelectedDiagram: vi.fn((onSelected) => {
    onSelected(null)
    return Object.assign(vi.fn(), { refresh: vi.fn() })
  }),
}))

afterEach(() => {
  cleanup()
  vi.unstubAllGlobals()
  vi.clearAllMocks()
})

it('shows a light code-only editor without a preview and does not insert automatically', async () => {
  vi.stubGlobal('Office', { onReady: vi.fn().mockResolvedValue({}), context: { document: {} } })
  const { container } = render(<PaneApp />)
  expect(await screen.findByRole('textbox', { name: 'Mermaid diagram source' })).toBeInTheDocument()
  expect(screen.getByRole('button', { name: 'Insert diagram' })).toBeInTheDocument()
  const footer = screen.getByRole('contentinfo')
  expect(within(footer).getByRole('button', { name: 'Insert diagram' })).toBeInTheDocument()
  expect(within(footer).getByRole('link', { name: 'Mermaid syntax' })).toBeInTheDocument()
  expect(container.querySelector('.editor-shell')?.nextElementSibling).toBe(footer)
  expect(within(screen.getByRole('banner')).queryByRole('link')).not.toBeInTheDocument()
  expect(screen.queryByRole('button', { name: 'Edit selected' })).not.toBeInTheDocument()
  expect(screen.queryByRole('button', { name: 'New diagram' })).not.toBeInTheDocument()
  expect(screen.queryByText(/Build .*Experimental pane/)).not.toBeInTheDocument()
  expect(screen.queryByText('Place the cursor in Word and insert your diagram to start live updates.'))
    .not.toBeInTheDocument()
  expect(screen.queryByRole('status')).not.toBeInTheDocument()
  expect(container.querySelector('.pane-build')).toBeNull()
  expect(screen.getByRole('combobox', { name: 'Diagram theme' })).toBeInTheDocument()
  expect(container.querySelector('.diagram-preview')).toBeNull()
  expect(screen.queryByRole('button', { name: 'Zoom in' })).not.toBeInTheDocument()
  expect(container.querySelector<HTMLElement>('.fui-FluentProvider')?.style.colorScheme).toBe('light')
  expect(insertDiagramWithPayload).not.toHaveBeenCalled()
  expect(updateDiagramById).not.toHaveBeenCalled()
})

it('hides Insert and saved-diagram instructions until the diagram is deselected', async () => {
  const existing = createDiagramPayload('flowchart LR\nA-->B', 'png', 'forest')
  vi.stubGlobal('Office', { onReady: vi.fn().mockResolvedValue({}), context: { document: {} } })
  vi.mocked(watchSelectedDiagram).mockImplementationOnce((onSelected) => {
    onSelected(existing)
    return Object.assign(vi.fn(), { refresh: vi.fn() })
  })
  render(<PaneApp />)
  await screen.findByRole('textbox', { name: 'Mermaid diagram source' })

  expect(screen.queryByRole('button', { name: 'Insert diagram' })).not.toBeInTheDocument()
  expect(screen.queryByText('Diagram is up to date in Word.')).not.toBeInTheDocument()
  expect(screen.queryByText('Select another diagram to edit it, or a blank line to insert a new one.'))
    .not.toBeInTheDocument()
  expect(screen.queryByRole('status')).not.toBeInTheDocument()
  expect(within(screen.getByRole('contentinfo')).getByRole('link', { name: 'Mermaid syntax' }))
    .toBeInTheDocument()

  act(() => {
    const [onSelected, , options] = vi.mocked(watchSelectedDiagram).mock.calls[0]
    options?.onSelectionChange?.()
    onSelected(null)
  })
  expect(within(screen.getByRole('contentinfo')).getByRole('button', { name: 'Insert diagram' }))
    .toBeInTheDocument()
})
