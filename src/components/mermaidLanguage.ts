import {
  defaultHighlightStyle,
  StreamLanguage,
  syntaxHighlighting,
  type StreamParser,
} from '@codemirror/language'

interface MermaidState {
  labelDepth: number
}

const diagramTypes = new Set([
  'architecture-beta',
  'block-beta',
  'c4component',
  'c4container',
  'c4context',
  'c4deployment',
  'classdiagram',
  'erdiagram',
  'flowchart',
  'gantt',
  'gitgraph',
  'graph',
  'journey',
  'kanban',
  'mindmap',
  'packet-beta',
  'pie',
  'quadrantchart',
  'requirementdiagram',
  'sankey-beta',
  'sequencediagram',
  'statediagram-v2',
  'timeline',
  'xychart-beta',
])

const keywords = new Set([
  'activate',
  'actor',
  'alt',
  'and',
  'autonumber',
  'class',
  'click',
  'critical',
  'deactivate',
  'direction',
  'else',
  'end',
  'linkstyle',
  'loop',
  'note',
  'opt',
  'par',
  'participant',
  'rect',
  'section',
  'style',
  'subgraph',
  'title',
])

const parser: StreamParser<MermaidState> = {
  name: 'mermaid',
  startState: () => ({ labelDepth: 0 }),
  token(stream, state) {
    if (stream.eatSpace()) {
      return null
    }
    if (stream.match(/%%.*$/)) {
      return 'comment'
    }
    if (stream.match(/"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'/)) {
      return 'string'
    }
    if (stream.match(/<?[-=.]+[ox]?[-=.]*>?/)) {
      return 'operator'
    }
    if (stream.match(/\d+(?:\.\d+)?/)) {
      return 'number'
    }
    if (stream.match('[') || stream.match('(') || stream.match('{')) {
      state.labelDepth += 1
      return 'bracket'
    }
    if (stream.match(/[\])}]/)) {
      state.labelDepth = Math.max(0, state.labelDepth - 1)
      return 'bracket'
    }
    if (stream.match(/[A-Za-z_][\w-]*/)) {
      const word = stream.current().toLowerCase()
      if (diagramTypes.has(word)) {
        return 'typeName'
      }
      if (keywords.has(word)) {
        return 'keyword'
      }
      if (word === 'true' || word === 'false') {
        return 'bool'
      }
      return state.labelDepth > 0 ? 'string' : 'variableName'
    }
    if (stream.match(/[|:;,]/)) {
      return 'punctuation'
    }

    stream.next()
    return null
  },
}

export const mermaidLanguage = [
  StreamLanguage.define(parser),
  syntaxHighlighting(defaultHighlightStyle),
]
