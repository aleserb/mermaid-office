import {
  snippetCompletion,
  type Completion,
  type CompletionContext,
  type CompletionResult,
} from '@codemirror/autocomplete'

const diagramSnippets: Completion[] = [
  snippetCompletion('flowchart ${LR}\n    ${A}[${Start}] --> ${B}[${Finish}]', {
    label: 'flowchart',
    detail: 'Flowchart template',
    type: 'keyword',
  }),
  snippetCompletion(
    'sequenceDiagram\n    participant ${Alice}\n    participant ${Bob}\n    ${Alice}->>${Bob}: ${Message}',
    {
      label: 'sequenceDiagram',
      detail: 'Sequence diagram template',
      type: 'keyword',
    },
  ),
  snippetCompletion(
    'classDiagram\n    class ${ClassName} {\n        +${method}()\n    }',
    {
      label: 'classDiagram',
      detail: 'Class diagram template',
      type: 'keyword',
    },
  ),
  snippetCompletion('stateDiagram-v2\n    [*] --> ${Active}\n    ${Active} --> [*]', {
    label: 'stateDiagram-v2',
    detail: 'State diagram template',
    type: 'keyword',
  }),
  snippetCompletion(
    'erDiagram\n    ${CUSTOMER} ||--o{ ${ORDER} : ${places}',
    {
      label: 'erDiagram',
      detail: 'Entity relationship template',
      type: 'keyword',
    },
  ),
  snippetCompletion('mindmap\n  root((${Topic}))\n    ${Branch}', {
    label: 'mindmap',
    detail: 'Mind map template',
    type: 'keyword',
  }),
]

const keywordOptions: Completion[] = [
  'subgraph',
  'end',
  'direction',
  'participant',
  'actor',
  'loop',
  'alt',
  'else',
  'opt',
  'par',
  'and',
  'critical',
  'note',
  'activate',
  'deactivate',
  'autonumber',
  'class',
  'style',
  'linkStyle',
  'click',
  'title',
  'section',
].map((label) => ({ label, type: 'keyword' }))

const identifier = '[A-Za-z_][\\w-]*'
const sequenceArrow = '(?:<<|<)?-{1,2}(?:>>|>|x|\\))'
const messageEndpoints = new RegExp(
  `^(${identifier}?)\\s*${sequenceArrow}[+-]?\\s*(${identifier})(?=\\s*:|\\s*$)`,
)
const messageRecipient = new RegExp(
  `^\\s*${identifier}?\\s*${sequenceArrow}[+-]?\\s*`,
)

function sequenceEntities(source: string, cursor: number): Completion[] {
  const entities = new Map<string, Completion>()
  let offset = 0
  for (const rawLine of source.split('\n')) {
    const line = rawLine.split('%%', 1)[0].trim()
    const declaration = /^(?:create\s+)?(participant|actor)\s+([A-Za-z_][\w-]*)(?:\s+as\s+(.+))?/.exec(line)
    if (declaration) {
      entities.set(declaration[2], {
        label: declaration[2],
        detail: declaration[3] ?? declaration[1],
        type: 'variable',
        boost: 10,
      })
    } else {
      const endpoints = messageEndpoints.exec(line)
      if (endpoints) {
        for (const id of endpoints.slice(1)) {
          // Don't turn the endpoint currently being typed into its own suggestion.
          if (cursor >= offset && cursor <= offset + rawLine.length && !entities.has(id)) {
            continue
          }
          if (!entities.has(id)) {
            entities.set(id, { label: id, detail: 'Participant', type: 'variable', boost: 10 })
          }
        }
      }
    }
    offset += rawLine.length + 1
  }
  return [...entities.values()]
}

export function mermaidCompletionSource(
  context: CompletionContext,
): CompletionResult | null {
  const source = context.state.doc.toString()
  const line = context.state.doc.lineAt(context.pos)
  const prefix = line.text.slice(0, context.pos - line.from)
  const isSequence = /^\s*sequenceDiagram\b/m.test(source)
  if (isSequence && (
    prefix.includes('%%') ||
    prefix.includes(':') ||
    /^\s*(?:create\s+)?(?:participant|actor)\s+\S+\s+as\s/.test(prefix) ||
    /^\s*(?:alt|else|loop|opt|par|and|critical|break|rect|box|title)\s/.test(prefix)
  )) {
    return null
  }

  const recipient = isSequence ? messageRecipient.exec(prefix) : null
  const referencePrefix = isSequence
    ? /^\s*(?:Note\s+(?:over\s+|(?:left|right)\s+of\s+)|(?:activate|deactivate|destroy)\s+)/i.exec(prefix)
    : null
  const referenceStart = recipient
    ? recipient[0].length
    : referencePrefix
      ? Math.max(referencePrefix[0].length, prefix.lastIndexOf(',') + 1)
      : 0
  const fragment = prefix.slice(referenceStart)
  const token = /[A-Za-z_][\w-]*$/.exec(fragment)
  const from = token ? context.pos - token[0].length : context.pos
  const reference = Boolean(recipient || referencePrefix)
  if (!token && !context.explicit && !reference) {
    return null
  }

  const atDiagramStart = line.number === 1 && prefix.slice(0, from - line.from).trim() === ''
  const entities = isSequence ? sequenceEntities(source, context.pos) : []

  return {
    from,
    options: atDiagramStart
      ? [...diagramSnippets, ...keywordOptions]
      : reference
        ? entities
        : [...entities, ...keywordOptions],
    validFor: /^[A-Za-z_][\w-]*$|^$/,
  }
}
