import type { Diagnostic } from '@codemirror/lint'
import type { Text } from '@codemirror/state'

export function createMermaidDiagnostic(
  document: Text,
  message: string,
): Diagnostic | null {
  if (!message) {
    return null
  }

  const lineMatch = message.match(/line\s+(\d+)/i)
  const lineNumber = Math.min(
    Math.max(Number.parseInt(lineMatch?.[1] ?? '1', 10), 1),
    document.lines,
  )
  const line = document.line(lineNumber)

  return {
    from: line.from,
    to: Math.max(line.from, line.to),
    severity: 'error',
    message,
  }
}
