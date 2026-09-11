export interface MermaidDiagnostic {
  message: string
  line: number
  column: number
  endColumn: number
}

interface ParserError {
  message?: unknown
  hash?: {
    token?: unknown
    expected?: unknown
    loc?: {
      first_line?: unknown
      first_column?: unknown
      last_column?: unknown
    }
  }
}

function asPositiveInteger(value: unknown, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value)
    ? Math.max(1, Math.trunc(value))
    : fallback
}

function normalizeToken(value: unknown): string | null {
  if (typeof value !== 'string' || !value || value === 'INVALID') {
    return null
  }
  return value === 'EOF' ? 'end of input' : value.replace(/^'|'$/g, '')
}

export function normalizeMermaidError(
  error: unknown,
  source = '',
): MermaidDiagnostic {
  const parserError =
    error && typeof error === 'object' ? (error as ParserError) : undefined
  const rawMessage =
    typeof parserError?.message === 'string'
      ? parserError.message
      : typeof error === 'string'
        ? error
        : 'Mermaid could not parse this diagram.'

  const rawLine = rawMessage.match(/line\s+(\d+)/i)?.[1]
  const caretLine = rawMessage.split('\n').find((line) => line.includes('^'))
  const loc = parserError?.hash?.loc
  const line = asPositiveInteger(
    loc?.first_line,
    rawLine ? Number.parseInt(rawLine, 10) : 1,
  )
  const caretColumn = caretLine?.indexOf('^')
  const token = normalizeToken(parserError?.hash?.token)
  const sourceLine = source.split(/\r?\n/)[line - 1]
  const column =
    token === 'end of input' && sourceLine !== undefined
      ? sourceLine.length + 1
      : typeof loc?.first_column === 'number'
      ? Math.max(1, Math.trunc(loc.first_column) + 1)
      : Math.max(1, (caretColumn ?? 0) + 1)
  const endColumn =
    typeof loc?.last_column === 'number'
      ? Math.max(column + 1, Math.trunc(loc.last_column) + 1)
      : column + 1
  const fallbackToken = normalizeToken(rawMessage.match(/got\s+'([^']+)'/i)?.[1])
  const unexpectedToken = token ?? fallbackToken
  const detail = unexpectedToken
    ? `Unexpected ${unexpectedToken}.`
    : rawMessage.split('\n')[0]?.replace(/^Error:\s*/i, '') ||
      'Mermaid could not parse this diagram.'

  return {
    line,
    column,
    endColumn,
    message: `Line ${line}, column ${column}: ${detail}`,
  }
}
