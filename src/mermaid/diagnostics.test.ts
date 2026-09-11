import { describe, expect, it } from 'vitest'
import { normalizeMermaidError } from './diagnostics'

describe('normalizeMermaidError', () => {
  it('extracts parser location and a concise message', () => {
    const diagnostic = normalizeMermaidError(
      {
        message: "Parse error on line 3:\nA -->\n-----^\nExpecting 'NODE', got 'EOF'",
        hash: {
          token: 'EOF',
          expected: ["'NODE'", "'TEXT'"],
          loc: { first_line: 2, first_column: 1, last_column: 0 },
        },
      },
      'flowchart LR\nA -->',
    )

    expect(diagnostic).toEqual({
      line: 2,
      column: 6,
      endColumn: 7,
      message: 'Line 2, column 6: Unexpected end of input.',
    })
  })
})
