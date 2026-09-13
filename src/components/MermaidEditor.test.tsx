import { cleanup, render } from '@testing-library/react'
import { undo } from '@codemirror/commands'
import { EditorView, keymap } from '@codemirror/view'
import { afterEach, expect, it, vi } from 'vitest'
import { MermaidEditor } from './MermaidEditor'

afterEach(() => {
  cleanup()
  vi.unstubAllGlobals()
})

it('highlights the active line and its number as the cursor moves', () => {
  const onChange = vi.fn()
  const { container, rerender } = render(
    <MermaidEditor value={'flowchart LR\nA-->B'} onChange={onChange} />,
  )
  const element = container.querySelector<HTMLElement>('.cm-editor')
  if (!element) throw new Error('Code editor was not rendered.')
  const editor = EditorView.findFromDOM(element)
  if (!editor) throw new Error('CodeMirror view was not found.')

  expect(container.querySelector('.cm-activeLine')).toHaveTextContent('flowchart LR')
  expect(container.querySelector('.cm-activeLineGutter')).toHaveTextContent('1')
  editor.dispatch({ selection: { anchor: editor.state.doc.line(2).from } })
  expect(container.querySelectorAll('.cm-activeLine')).toHaveLength(1)
  expect(container.querySelector('.cm-activeLine')).toHaveTextContent('A-->B')
  expect(container.querySelector('.cm-activeLineGutter')).toHaveTextContent('2')
  expect(onChange).not.toHaveBeenCalled()

  rerender(<MermaidEditor value="sequenceDiagram" historyKey="another-diagram" onChange={onChange} />)
  expect(container.querySelector('.cm-activeLine')).toHaveTextContent('sequenceDiagram')
  expect(container.querySelector('.cm-activeLineGutter')).toHaveTextContent('1')
})

it('does not register editor search shortcuts or open a search and replace panel', () => {
  const { container } = render(<MermaidEditor value="flowchart LR" onChange={vi.fn()} />)
  const element = container.querySelector<HTMLElement>('.cm-editor')
  if (!element) throw new Error('Code editor was not rendered.')
  const editor = EditorView.findFromDOM(element)
  if (!editor) throw new Error('CodeMirror view was not found.')
  const bindings = editor.state.facet(keymap).flat()
  for (const key of ['Mod-f', 'F3', 'Mod-g', 'Mod-Shift-l', 'Mod-Alt-g', 'Mod-d']) {
    expect(bindings.some((binding) => binding.key === key)).toBe(false)
  }
  for (const modifiers of [{ ctrlKey: true }, { metaKey: true }]) {
    editor.contentDOM.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'f', code: 'KeyF', bubbles: true, cancelable: true, ...modifiers,
    }))
  }
  expect(container.querySelector('.cm-search')).toBeNull()
  expect(container.querySelector('.cm-panel')).toBeNull()
})

it('uses System UI at 12px and preserves editor state, selection, and undo history', () => {
  const source = 'flowchart LR\nA-->B'
  const onChange = vi.fn()
  const { container, rerender } = render(<MermaidEditor value={source} onChange={onChange} />)
  const element = container.querySelector<HTMLElement>('.cm-editor')
  if (!element) throw new Error('Code editor was not rendered.')
  const editor = EditorView.findFromDOM(element)
  if (!editor) throw new Error('CodeMirror view was not found.')
  const host = container.querySelector('.editor')
  expect(host).toHaveStyle({
    fontFamily: 'system-ui, sans-serif',
    fontSize: '12px',
  })
  editor.dispatch({
    changes: { from: source.length, insert: '\nB-->C' },
    selection: { anchor: source.length },
  })
  const editedSource = editor.state.doc.toString()
  rerender(<MermaidEditor value={editedSource} onChange={onChange} />)
  const state = editor.state
  onChange.mockClear()

  rerender(
    <MermaidEditor
      value={editedSource}
      onChange={onChange}
    />,
  )
  expect(EditorView.findFromDOM(element)).toBe(editor)
  expect(editor.state).toBe(state)
  expect(editor.state.selection.main.anchor).toBe(source.length)
  expect(host).toHaveStyle({ fontFamily: 'system-ui, sans-serif', fontSize: '12px' })
  expect(onChange).not.toHaveBeenCalled()
  expect(undo(editor)).toBe(true)
  expect(editor.state.doc.toString()).toBe(source)
})

it('switches between light and dark themes without losing editor state or history', () => {
  const source = 'sequenceDiagram\nparticipant web'
  const onChange = vi.fn()
  const { container, rerender } = render(
    <MermaidEditor value={source} onChange={onChange} darkMode={false} />,
  )
  const element = container.querySelector<HTMLElement>('.cm-editor')
  if (!element) throw new Error('Code editor was not rendered.')
  const editor = EditorView.findFromDOM(element)
  if (!editor) throw new Error('CodeMirror view was not found.')
  expect(editor.state.facet(EditorView.darkTheme)).toBe(false)
  editor.dispatch({
    changes: { from: source.length, insert: '\nweb->>db: Request' },
    selection: { anchor: source.length },
  })
  const editedSource = editor.state.doc.toString()

  rerender(<MermaidEditor value={editedSource} onChange={onChange} darkMode />)
  expect(editor.state.facet(EditorView.darkTheme)).toBe(true)
  expect(editor.state.doc.toString()).toBe(editedSource)
  expect(editor.state.selection.main.anchor).toBe(source.length)
  expect(undo(editor)).toBe(true)
  expect(editor.state.doc.toString()).toBe(source)

  rerender(<MermaidEditor value={source} onChange={onChange} darkMode={false} />)
  expect(editor.state.facet(EditorView.darkTheme)).toBe(false)
})
