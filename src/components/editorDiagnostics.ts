import type { Diagnostic } from '@codemirror/lint'
import type { Text } from '@codemirror/state'
import type { MermaidDiagnostic } from '../mermaid/diagnostics'

export function createMermaidDiagnostic(
  document: Text,
  diagnostic: MermaidDiagnostic | null,
): Diagnostic | null {
  if (!diagnostic) {
    return null
  }

  const lineNumber = Math.min(
    Math.max(diagnostic.line, 1),
    document.lines,
  )
  const line = document.line(lineNumber)
  const from = Math.min(line.to, line.from + Math.max(0, diagnostic.column - 1))
  const to = Math.min(
    line.to,
    Math.max(from + 1, line.from + Math.max(diagnostic.endColumn - 1, 1)),
  )

  return {
    from,
    to,
    severity: 'error',
    message: diagnostic.message,
  }
}
