import {
  HighlightStyle,
  StreamLanguage,
  syntaxHighlighting,
  type StreamParser,
} from '@codemirror/language'
import { tags } from '@lezer/highlight'

interface MermaidState {
  labelDepth: number
  restOfLineIsText: boolean
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
  'as',
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
  'over',
  'par',
  'participant',
  'rect',
  'section',
  'style',
  'subgraph',
  'title',
])

const textFollowingKeywords = new Set([
  'alt',
  'and',
  'else',
  'loop',
  'opt',
  'par',
  'rect',
  'section',
  'title',
])

const parser: StreamParser<MermaidState> = {
  name: 'mermaid',
  startState: () => ({ labelDepth: 0, restOfLineIsText: false }),
  token(stream, state) {
    if (stream.sol()) {
      state.labelDepth = 0
      state.restOfLineIsText = false
    }
    if (stream.eatSpace()) {
      return null
    }
    if (state.restOfLineIsText) {
      stream.skipToEnd()
      return 'string'
    }
    if (stream.match(/%%.*$/)) {
      return 'comment'
    }
    if (stream.match(/"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'/)) {
      return 'string'
    }
    if (stream.match(/(?:<<|<)?[-=.]+(?:>>|>|[ox]|\))?[+-]?/)) {
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
    if (stream.match(/[A-Za-z_]\w*(?:-[A-Za-z_]\w*)*/)) {
      const word = stream.current().toLowerCase()
      if (diagramTypes.has(word)) {
        return 'typeName'
      }
      if (keywords.has(word)) {
        if (word === 'as' || textFollowingKeywords.has(word)) {
          state.restOfLineIsText = true
        }
        return 'keyword'
      }
      if (word === 'true' || word === 'false') {
        return 'bool'
      }
      return state.labelDepth > 0 ? 'string' : 'variableName'
    }
    if (stream.match(':')) {
      state.restOfLineIsText = true
      return 'punctuation'
    }
    if (stream.match(/[|;,]/)) {
      return 'punctuation'
    }

    stream.next()
    return null
  },
}

const mermaidHighlightStyle = HighlightStyle.define([
  { tag: tags.typeName, color: '#008855', fontWeight: '600' },
  { tag: tags.keyword, color: '#770088', fontWeight: '600' },
  { tag: tags.variableName, color: '#005a9e' },
  { tag: tags.string, color: '#a31515' },
  { tag: tags.operator, color: '#005fb8', fontWeight: '600' },
  { tag: tags.comment, color: '#008000', fontStyle: 'italic' },
  { tag: tags.number, color: '#098658' },
  { tag: [tags.bracket, tags.punctuation], color: '#5c2d91' },
], { themeType: 'light' })

export const mermaidLanguage = [
  StreamLanguage.define(parser),
  syntaxHighlighting(mermaidHighlightStyle),
]
