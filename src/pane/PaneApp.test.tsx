import { act, cleanup, render, screen, waitFor, within } from '@testing-library/react'
import { EditorView } from '@codemirror/view'
import userEvent from '@testing-library/user-event'
import { afterEach, expect, it, vi } from 'vitest'
import { PaneApp } from './PaneApp'
import { insertDiagramWithPayload, updateDiagramById } from '../word/insertDiagram'
import { createDiagramPayload } from '../metadata/payload'
import { DEFAULT_DIAGRAM_SETTINGS } from '../metadata/diagramSettings'
import { watchSelectedDiagram } from '../word/selection'
import { renderMermaid } from '../mermaid/render'

vi.mock('../mermaid/render', async (importOriginal) => ({
  ...await importOriginal<typeof import('../mermaid/render')>(),
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
  vi.restoreAllMocks()
  vi.clearAllMocks()
  localStorage.clear()
})

it('uses fixed System UI at 12px without font or size controls', async () => {
  vi.stubGlobal('Office', { onReady: vi.fn().mockResolvedValue({}), context: { document: {} } })
  const { container } = render(<PaneApp />)
  await screen.findByRole('textbox', { name: 'Mermaid diagram source' })
  expect(screen.queryByRole('group', { name: 'Code editor display' })).not.toBeInTheDocument()
  expect(screen.queryByRole('combobox', { name: 'Code editor font' })).not.toBeInTheDocument()
  expect(screen.queryByRole('combobox', { name: 'Code editor font size' })).not.toBeInTheDocument()
  expect(container.querySelector('.editor')).toHaveStyle({
    fontFamily: 'system-ui, sans-serif',
    fontSize: '12px',
  })
  expect(insertDiagramWithPayload).not.toHaveBeenCalled()
  expect(updateDiagramById).not.toHaveBeenCalled()
})

it('shows a light code-only editor without a preview and does not insert automatically', async () => {
  vi.stubGlobal('Office', { onReady: vi.fn().mockResolvedValue({}), context: { document: {} } })
  const { container } = render(<PaneApp />)
  expect(await screen.findByRole('textbox', { name: 'Mermaid diagram source' })).toBeInTheDocument()
  expect(screen.getByRole('button', { name: 'Insert' })).toBeInTheDocument()
  const footer = screen.getByRole('contentinfo')
  expect(within(footer).getByRole('button', { name: 'Insert' })).toBeInTheDocument()
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
  expect(screen.queryByRole('combobox', { name: 'Diagram theme' })).not.toBeInTheDocument()
  expect(within(screen.getByRole('banner')).getByRole('button', { name: 'Diagram settings' }))
    .toHaveAttribute('title', 'Diagram settings')
  expect(container.querySelector('.diagram-preview')).toBeNull()
  expect(screen.queryByRole('button', { name: 'Zoom in' })).not.toBeInTheDocument()
  expect(container.querySelector<HTMLElement>('.fui-FluentProvider')?.style.colorScheme).toBe('light')
  expect(insertDiagramWithPayload).not.toHaveBeenCalled()
  expect(updateDiagramById).not.toHaveBeenCalled()
})

it('opens settings from the keyboard and discards edits on Escape and Cancel', async () => {
  vi.stubGlobal('Office', { onReady: vi.fn().mockResolvedValue({}), context: { document: {} } })
  const user = userEvent.setup()
  render(<PaneApp />)
  await screen.findByRole('textbox', { name: 'Mermaid diagram source' })
  const gear = screen.getByRole('button', { name: 'Diagram settings' })
  gear.focus()
  await user.keyboard('{Enter}')
  expect(await screen.findByRole('dialog', { name: 'Diagram settings' })).toBeInTheDocument()
  const originalTheme = (screen.getByRole('combobox', { name: 'Diagram theme' }) as HTMLSelectElement).value
  await user.selectOptions(screen.getByRole('combobox', { name: 'Diagram theme' }), 'dark')
  await user.keyboard('{Escape}')
  expect(screen.queryByRole('dialog')).not.toBeInTheDocument()
  await waitFor(() => expect(gear).toHaveFocus())
  await user.click(gear)
  expect(screen.getByRole('combobox', { name: 'Diagram theme' })).toHaveValue(originalTheme)
  await user.click(screen.getByRole('button', { name: 'Cancel' }))
  expect(screen.queryByRole('dialog')).not.toBeInTheDocument()
  expect(insertDiagramWithPayload).not.toHaveBeenCalled()
  expect(updateDiagramById).not.toHaveBeenCalled()
})

it('closes stale settings when the selected diagram changes and opens the new values', async () => {
  const first = createDiagramPayload('flowchart LR\nA-->B', 'png', 'forest')
  const second = createDiagramPayload('flowchart TD\nC-->D', 'png', 'neutral')
  vi.stubGlobal('Office', { onReady: vi.fn().mockResolvedValue({}), context: { document: {} } })
  vi.mocked(watchSelectedDiagram).mockImplementationOnce(onSelected => {
    onSelected(first)
    return Object.assign(vi.fn(), { refresh: vi.fn() })
  })
  const user = userEvent.setup()
  render(<PaneApp />)
  await screen.findByRole('textbox', { name: 'Mermaid diagram source' })
  await user.click(screen.getByRole('button', { name: 'Diagram settings' }))
  await user.selectOptions(await screen.findByRole('combobox', { name: 'Diagram theme' }), 'dark')
  act(() => {
    const [onSelected, , options] = vi.mocked(watchSelectedDiagram).mock.calls[0]
    options?.onSelectionChange?.()
    onSelected(second)
  })
  expect(screen.queryByRole('dialog')).not.toBeInTheDocument()
  expect(updateDiagramById).not.toHaveBeenCalled()
  await user.click(screen.getByRole('button', { name: 'Diagram settings' }))
  expect(await screen.findByRole('combobox', { name: 'Diagram theme' })).toHaveValue('neutral')
})

it('renders all settings together only after Apply and retains them on reopening', async () => {
  vi.stubGlobal('Office', { onReady: vi.fn().mockResolvedValue({}), context: { document: {} } })
  const user = userEvent.setup()
  render(<PaneApp />)
  await waitFor(() => expect(renderMermaid).toHaveBeenCalled(), { timeout: 2000 })
  vi.mocked(renderMermaid).mockClear()
  await user.click(screen.getByRole('button', { name: 'Diagram settings' }))
  await user.selectOptions(screen.getByRole('combobox', { name: 'Diagram theme' }), 'forest')
  await user.selectOptions(screen.getByRole('combobox', { name: 'Font size' }), '20')
  await user.selectOptions(screen.getByRole('combobox', { name: 'Image quality' }), 'high')
  expect(renderMermaid).not.toHaveBeenCalled()
  expect(updateDiagramById).not.toHaveBeenCalled()
  await user.click(screen.getByRole('button', { name: 'Apply' }))
  expect(screen.queryByRole('dialog')).not.toBeInTheDocument()
  await waitFor(() => expect(renderMermaid).toHaveBeenCalledExactlyOnceWith(
    expect.any(String), 'forest', { ...DEFAULT_DIAGRAM_SETTINGS, fontSize: 20, imageQuality: 'high' },
  ), { timeout: 2000 })
  await user.click(screen.getByRole('button', { name: 'Diagram settings' }))
  expect(screen.getByRole('combobox', { name: 'Diagram theme' })).toHaveValue('forest')
  expect(screen.getByRole('combobox', { name: 'Font size' })).toHaveValue('20')
  expect(screen.getByRole('combobox', { name: 'Image quality' })).toHaveValue('high')
  expect(insertDiagramWithPayload).not.toHaveBeenCalled()
})

it('uses the selected diagram header for conditional settings', async () => {
  const selected = createDiagramPayload(
    '---\ntitle: flowchart example\n---\n%% flowchart comment\nsequenceDiagram\nAlice->>Bob: Hello',
    'png', 'default',
  )
  vi.stubGlobal('Office', { onReady: vi.fn().mockResolvedValue({}), context: { document: {} } })
  vi.mocked(watchSelectedDiagram).mockImplementationOnce(onSelected => {
    onSelected(selected)
    return Object.assign(vi.fn(), { refresh: vi.fn() })
  })
  const user = userEvent.setup()
  render(<PaneApp />)
  await screen.findByRole('textbox', { name: 'Mermaid diagram source' })
  await user.click(screen.getByRole('button', { name: 'Diagram settings' }))
  expect(screen.getByRole('checkbox', { name: 'Number messages' })).toBeInTheDocument()
  expect(screen.getByRole('checkbox', { name: 'Wrap text' })).toBeInTheDocument()
  expect(screen.getByRole('checkbox', { name: 'Repeat participants below' })).toBeInTheDocument()
  expect(screen.getByRole('combobox', { name: 'Spacing' })).toBeInTheDocument()
  expect(screen.queryByRole('combobox', { name: 'Connector style' })).not.toBeInTheDocument()
  expect(screen.queryByRole('combobox', { name: 'Layout' })).not.toBeInTheDocument()
})

it('hides Insert and saved-diagram instructions until the diagram is deselected', async () => {
  vi.spyOn(document, 'hasFocus').mockReturnValue(false)
  const existing = createDiagramPayload('flowchart LR\nA-->B', 'png', 'forest')
  vi.stubGlobal('Office', { onReady: vi.fn().mockResolvedValue({}), context: { document: {} } })
  vi.mocked(watchSelectedDiagram).mockImplementationOnce((onSelected) => {
    onSelected(existing)
    return Object.assign(vi.fn(), { refresh: vi.fn() })
  })

  render(<PaneApp />)
  await screen.findByRole('textbox', { name: 'Mermaid diagram source' })

  expect(screen.queryByRole('button', { name: 'Insert' })).not.toBeInTheDocument()
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
  expect(within(screen.getByRole('contentinfo')).getByRole('button', { name: 'Insert' }))
    .toBeInTheDocument()
})

it('replaces settings with the pending-edits confirmation instead of stacking modals', async () => {
  vi.stubGlobal('Office', { onReady: vi.fn().mockResolvedValue({}), context: { document: {} } })
  const user = userEvent.setup()
  const { container } = render(<PaneApp />)
  await screen.findByRole('textbox', { name: 'Mermaid diagram source' })
  const code = EditorView.findFromDOM(container.querySelector<HTMLElement>('.cm-editor')!)
  act(() => code!.dispatch({ changes: { from: code!.state.doc.length, insert: '\nB-->C' } }))
  await user.click(screen.getByRole('button', { name: 'Diagram settings' }))
  await user.selectOptions(screen.getByRole('combobox', { name: 'Diagram theme' }), 'dark')
  act(() => {
    const [onSelected] = vi.mocked(watchSelectedDiagram).mock.calls[0]
    onSelected(createDiagramPayload('flowchart LR\nX-->Y', 'png', 'forest'))
  })
  expect(screen.getAllByRole('dialog')).toHaveLength(1)
  expect(screen.getByRole('dialog', { name: 'Discard pending edits?' })).toBeInTheDocument()
  expect(screen.queryByRole('dialog', { name: 'Diagram settings' })).not.toBeInTheDocument()
  await user.click(screen.getByRole('button', { name: 'Keep editing' }))
  expect(screen.queryByRole('dialog')).not.toBeInTheDocument()
  expect(screen.getByRole('button', { name: 'Diagram settings' })).toBeEnabled()
  expect(updateDiagramById).not.toHaveBeenCalled()
})

it('disables settings while loading the selected diagram', async () => {
  vi.stubGlobal('Office', { onReady: vi.fn().mockResolvedValue({}), context: { document: {} } })
  const user = userEvent.setup()
  render(<PaneApp />)
  expect(screen.getByRole('button', { name: 'Diagram settings' })).toBeDisabled()
  await screen.findByRole('textbox', { name: 'Mermaid diagram source' })
  await user.click(screen.getByRole('button', { name: 'Diagram settings' }))
  act(() => {
    const [, , options] = vi.mocked(watchSelectedDiagram).mock.calls[0]
    options?.onSelectionChange?.()
  })
  expect(screen.getByRole('button', { name: 'Apply' })).toBeDisabled()
  await user.click(screen.getByRole('button', { name: 'Cancel' }))
  expect(screen.getByRole('button', { name: 'Diagram settings' })).toBeDisabled()
})

it.each(['render', 'word'])('shows %s errors below the code editor and above the footer', async (kind) => {
  vi.stubGlobal('Office', { onReady: vi.fn().mockResolvedValue({}), context: { document: {} } })
  vi.stubGlobal('ResizeObserver', class {
    observe() {}
    unobserve() {}
    disconnect() {}
  })
  const message = kind === 'render' ? 'Invalid diagram source.' : 'Unable to read Word selection.'
  if (kind === 'render') {
    vi.mocked(renderMermaid).mockRejectedValueOnce(new Error(message))
  } else {
    vi.mocked(watchSelectedDiagram).mockImplementationOnce((_onSelected, onError) => {
      onError(new Error(message))
      return Object.assign(vi.fn(), { refresh: vi.fn() })
    })
  }
  const { container } = render(<PaneApp />)
  const error = await screen.findByText(message, { exact: false }, { timeout: 2000 })
  const bar = error.closest('.fui-MessageBar')
  expect(container.querySelector('.editor-shell')?.nextElementSibling).toBe(bar)
  expect(bar?.nextElementSibling).toBe(screen.getByRole('contentinfo'))
  expect(screen.getByRole('button', { name: 'Insert' })).toBeDisabled()
})
