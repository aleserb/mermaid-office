import { EditorState } from '@codemirror/state'
import { describe, expect, it } from 'vitest'
import { createMermaidDiagnostic } from './editorDiagnostics'

describe('createMermaidDiagnostic', () => {
  it('targets the line reported by Mermaid', () => {
    const document = EditorState.create({ doc: 'flowchart LR\nA -->\nB' }).doc
    const diagnostic = createMermaidDiagnostic(document, 'Parse error on line 2')

    expect(diagnostic).toMatchObject({
      from: document.line(2).from,
      to: document.line(2).to,
      severity: 'error',
    })
  })

  it('falls back to the first line', () => {
    const document = EditorState.create({ doc: 'invalid' }).doc
    expect(createMermaidDiagnostic(document, 'Syntax error')?.from).toBe(0)
  })
})
