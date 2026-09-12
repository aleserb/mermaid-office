import { cleanup, render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { FluentProvider, webLightTheme } from '@fluentui/react-components'
import { afterEach, expect, it, vi } from 'vitest'
import { DEFAULT_DIAGRAM_SETTINGS, type DiagramKind } from '../metadata/diagramSettings'
import { DiagramSettingsDialog } from './DiagramSettingsDialog'

afterEach(cleanup)

function setup(diagramKind: DiagramKind = 'flowchart', disabled = false) {
  const onApply = vi.fn()
  const onCancel = vi.fn()
  const view = render(
    <FluentProvider theme={webLightTheme}>
      <DiagramSettingsDialog theme="default" settings={DEFAULT_DIAGRAM_SETTINGS}
        diagramKind={diagramKind} disabled={disabled} onApply={onApply} onCancel={onCancel} />
    </FluentProvider>,
  )
  return { ...view, onApply, onCancel, user: userEvent.setup() }
}

it('offers diagram typography, themes, looks, flowchart options and PNG quality', () => {
  setup()
  expect(screen.getByRole('dialog', { name: 'Diagram settings' })).toHaveAttribute('aria-modal', 'true')
  for (const name of ['Diagram theme', 'Font family', 'Font size', 'Visual look', 'Spacing', 'Connector style', 'Layout', 'Image quality']) {
    expect(screen.getByRole('combobox', { name })).toBeInTheDocument()
  }
  expect(screen.queryByRole('checkbox')).not.toBeInTheDocument()
  expect(screen.queryByLabelText(/custom color/i)).not.toBeInTheDocument()
  expect(screen.getByText(/output PNG, not the editor font/)).toBeInTheDocument()
  expect(screen.getByText(/support varies by diagram type/)).toBeInTheDocument()
  expect(screen.getByRole('combobox', { name: 'Image quality' })).toHaveValue('auto')
})

it('applies one atomic copy without changing caller defaults or applying control edits live', async () => {
  const { user, onApply } = setup()
  for (const [name, value] of [
    ['Diagram theme', 'forest'], ['Font family', 'Arial'], ['Font size', '18'],
    ['Visual look', 'handDrawn'], ['Spacing', 'compact'], ['Connector style', 'basis'],
    ['Layout', 'elk'], ['Image quality', 'high'],
  ]) {
    await user.selectOptions(screen.getByRole('combobox', { name }), value)
  }
  expect(onApply).not.toHaveBeenCalled()
  expect(DEFAULT_DIAGRAM_SETTINGS.fontFamily).toBe('theme')
  expect(DEFAULT_DIAGRAM_SETTINGS.imageQuality).toBe('auto')
  expect(screen.getByText(/larger, slower PNGs.*Word's loading dialog/)).toBeInTheDocument()
  await user.click(screen.getByRole('button', { name: 'Apply' }))
  expect(onApply).toHaveBeenCalledExactlyOnceWith('forest', {
    ...DEFAULT_DIAGRAM_SETTINGS,
    fontFamily: 'Arial', fontSize: 18, look: 'handDrawn', spacing: 'compact',
    curve: 'basis', layout: 'elk', imageQuality: 'high',
  })
  expect(onApply.mock.calls[0][1]).not.toBe(DEFAULT_DIAGRAM_SETTINGS)
})

it.each(['Cancel', 'Escape'])('discards local edits with %s', async action => {
  const { user, onApply, onCancel } = setup()
  await user.selectOptions(screen.getByRole('combobox', { name: 'Diagram theme' }), 'dark')
  if (action === 'Cancel') await user.click(screen.getByRole('button', { name: 'Cancel' }))
  else await user.keyboard('{Escape}')
  expect(onCancel).toHaveBeenCalledOnce()
  expect(onApply).not.toHaveBeenCalled()
})

it('shows and applies sequence controls without flowchart-only controls', async () => {
  const { user, onApply } = setup('sequence')
  expect(screen.getByRole('option', { name: 'Hand-drawn' })).toBeDisabled()
  expect(screen.getByText(/Hand-drawn is not supported for sequence diagrams/)).toBeInTheDocument()
  expect(screen.getByRole('combobox', { name: 'Spacing' })).toBeInTheDocument()
  expect(screen.queryByRole('combobox', { name: 'Connector style' })).not.toBeInTheDocument()
  expect(screen.queryByRole('combobox', { name: 'Layout' })).not.toBeInTheDocument()
  expect(screen.getByRole('checkbox', { name: 'Repeat participants below' })).toBeChecked()
  for (const name of ['Number messages', 'Wrap text', 'Repeat participants below']) {
    await user.click(screen.getByRole('checkbox', { name }))
  }
  await user.click(screen.getByRole('button', { name: 'Apply' }))
  expect(onApply).toHaveBeenCalledExactlyOnceWith('default', {
    ...DEFAULT_DIAGRAM_SETTINGS,
    sequenceNumbers: true, sequenceWrap: true, sequenceMirrorActors: false,
  })
})

it('preserves a saved Hand-drawn sequence choice instead of silently replacing it', async () => {
  const onApply = vi.fn()
  const user = userEvent.setup()
  render(<DiagramSettingsDialog theme="default"
    settings={{ ...DEFAULT_DIAGRAM_SETTINGS, look: 'handDrawn' }}
    diagramKind="sequence" onApply={onApply} onCancel={vi.fn()} />)
  expect(screen.getByRole('combobox', { name: 'Visual look' })).toHaveValue('handDrawn')
  expect(screen.getByRole('option', { name: 'Hand-drawn' })).toBeDisabled()
  await user.click(screen.getByRole('button', { name: 'Apply' }))
  expect(onApply).toHaveBeenCalledExactlyOnceWith('default', {
    ...DEFAULT_DIAGRAM_SETTINGS, look: 'handDrawn',
  })
})

it.each<DiagramKind>(['class', 'er', 'state', 'other'])('hides unsupported controls for %s diagrams', kind => {
  setup(kind)
  for (const name of ['Spacing', 'Connector style', 'Layout']) {
    expect(screen.queryByRole('combobox', { name })).not.toBeInTheDocument()
  }
  expect(screen.queryByRole('checkbox')).not.toBeInTheDocument()
  expect(screen.getByRole('combobox', { name: 'Visual look' })).toBeInTheDocument()
})

it('disables Apply while a document operation is in progress but permits Cancel', async () => {
  const { user, onApply, onCancel } = setup('flowchart', true)
  expect(screen.getByRole('button', { name: 'Apply' })).toBeDisabled()
  await user.click(screen.getByRole('button', { name: 'Apply' }))
  expect(onApply).not.toHaveBeenCalled()
  await user.click(screen.getByRole('button', { name: 'Cancel' }))
  expect(onCancel).toHaveBeenCalledOnce()
})

it('initializes fresh values when reopened and preserves hidden type-specific settings', async () => {
  const { user, unmount, onApply } = setup()
  await user.selectOptions(screen.getByRole('combobox', { name: 'Diagram theme' }), 'dark')
  unmount()
  render(<DiagramSettingsDialog theme="forest" settings={{
    ...DEFAULT_DIAGRAM_SETTINGS, imageQuality: 'standard', layout: 'elk', sequenceWrap: true,
  }} diagramKind="other" onCancel={vi.fn()} onApply={onApply} />)
  expect(screen.getByRole('combobox', { name: 'Diagram theme' })).toHaveValue('forest')
  expect(screen.getByRole('combobox', { name: 'Image quality' })).toHaveValue('standard')
  await user.click(screen.getByRole('button', { name: 'Apply' }))
  expect(onApply).toHaveBeenCalledExactlyOnceWith('forest', {
    ...DEFAULT_DIAGRAM_SETTINGS, imageQuality: 'standard', layout: 'elk', sequenceWrap: true,
  })
})

it('restores background accessibility on controlled close and resets a quickly reopened form', async () => {
  const user = userEvent.setup()
  const onApply = vi.fn()
  const content = (open: boolean) => (
    <FluentProvider theme={webLightTheme}>
      <button>Background editor</button>
      <DiagramSettingsDialog open={open} theme="default" settings={DEFAULT_DIAGRAM_SETTINGS}
        diagramKind="flowchart" onApply={onApply} onCancel={vi.fn()} />
    </FluentProvider>
  )
  const { rerender } = render(content(false))
  rerender(content(true))
  await user.selectOptions(screen.getByRole('combobox', { name: 'Image quality' }), 'high')
  rerender(content(false))
  expect(await screen.findByRole('button', { name: 'Background editor' })).toBeVisible()
  rerender(content(true))
  expect(screen.getByRole('combobox', { name: 'Image quality' })).toHaveValue('auto')
  await user.click(screen.getByRole('button', { name: 'Apply' }))
  expect(onApply).toHaveBeenCalledExactlyOnceWith('default', DEFAULT_DIAGRAM_SETTINGS)
})
