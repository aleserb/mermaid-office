import { EditorState } from '@codemirror/state'
import { EditorView } from '@codemirror/view'
import { afterEach, describe, expect, it } from 'vitest'
import { mermaidLanguage } from './mermaidLanguage'

let editor: EditorView | undefined

afterEach(() => {
  editor?.destroy()
  editor = undefined
  document.body.replaceChildren()
})

function highlightedTokens(source: string): string[] {
  const parent = document.createElement('div')
  document.body.append(parent)
  editor = new EditorView({
    parent,
    state: EditorState.create({
      doc: source,
      extensions: mermaidLanguage,
    }),
  })

  return Array.from(parent.querySelectorAll('.cm-content span'))
    .map((element) => element.textContent ?? '')
    .filter(Boolean)
}

describe('Mermaid syntax highlighting', () => {
  it('highlights sequence aliases, messages, and arrow variants', () => {
    const tokens = highlightedTokens(`sequenceDiagram
actor web as Web Browser
Note over web,db: Signed in
web->>+account: Login
account->>-web: Complete
blog--)mail: Notify
blog-->>-web: Posted
alt Credentials found
end`)

    expect(tokens).toEqual(
      expect.arrayContaining([
        'sequenceDiagram',
        'actor',
        'as',
        'Web Browser',
        'Note',
        'over',
        '->>+',
        '->>-',
        '--)',
        '-->>-',
        'Credentials found',
      ]),
    )
  })

  it('keeps flowchart circle and cross edges intact', () => {
    const tokens = highlightedTokens(
      'flowchart LR\nA --o B\nB --x C\nC <-> D\nD <--> E\nE <-x F\nF <--x G',
    )

    expect(tokens).toEqual(
      expect.arrayContaining(['--o', '--x', '<->', '<-->', '<-x', '<--x']),
    )
  })
})
