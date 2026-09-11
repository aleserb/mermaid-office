import { EditorState } from '@codemirror/state'
import { describe, expect, it } from 'vitest'
import { createMermaidDiagnostic } from './editorDiagnostics'

describe('createMermaidDiagnostic', () => {
  it('targets the line reported by Mermaid', () => {
    const document = EditorState.create({ doc: 'flowchart LR\nA -->\nB' }).doc
    const diagnostic = createMermaidDiagnostic(document, {
      message: 'Unexpected end of input.',
      line: 2,
      column: 3,
      endColumn: 4,
    })

    expect(diagnostic).toMatchObject({
      from: document.line(2).from + 2,
      to: document.line(2).from + 3,
      severity: 'error',
    })
  })

  it('falls back to the first line', () => {
    const document = EditorState.create({ doc: 'invalid' }).doc
    expect(
      createMermaidDiagnostic(document, {
        message: 'Syntax error',
        line: 1,
        column: 1,
        endColumn: 2,
      })?.from,
    ).toBe(0)
  })
})
