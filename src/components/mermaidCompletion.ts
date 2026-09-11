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

export function mermaidCompletionSource(
  context: CompletionContext,
): CompletionResult | null {
  const word = context.matchBefore(/[A-Za-z_-]*/)
  if (!word || (!context.explicit && word.from === word.to)) {
    return null
  }

  const line = context.state.doc.lineAt(context.pos)
  const atDiagramStart = line.number === 1 && line.text.slice(0, word.from).trim() === ''

  return {
    from: word.from,
    options: atDiagramStart
      ? [...diagramSnippets, ...keywordOptions]
      : keywordOptions,
    validFor: /^[A-Za-z_-]*$/,
  }
}
