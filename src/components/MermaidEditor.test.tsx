import { cleanup, render } from '@testing-library/react'
import { undo } from '@codemirror/commands'
import { EditorView } from '@codemirror/view'
import { afterEach, expect, it, vi } from 'vitest'
import { MermaidEditor } from './MermaidEditor'

afterEach(() => {
  cleanup()
  vi.unstubAllGlobals()
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

it('stays light under dark Word and system preferences and preserves editing history', () => {
  vi.stubGlobal('Office', { context: { officeTheme: { isDarkTheme: true } } })
  vi.stubGlobal('matchMedia', vi.fn().mockReturnValue({
    matches: true,
    addEventListener: vi.fn(),
    removeEventListener: vi.fn(),
    addListener: vi.fn(),
    removeListener: vi.fn(),
  }))
  const source = 'sequenceDiagram\nparticipant web'
  const onChange = vi.fn()
  const { container, rerender } = render(
    <MermaidEditor value={source} onChange={onChange} />,
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

  rerender(<MermaidEditor value={source} onChange={onChange} />)
  expect(editor.state.facet(EditorView.darkTheme)).toBe(false)
  expect(editor.state.doc.toString()).toBe(editedSource)
  expect(editor.state.selection.main.anchor).toBe(source.length)
  expect(undo(editor)).toBe(true)
  expect(editor.state.doc.toString()).toBe(source)
})
