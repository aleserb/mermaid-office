import {
  acceptCompletion,
  autocompletion,
  CompletionContext,
  currentCompletions,
  startCompletion,
} from '@codemirror/autocomplete'
import { EditorState } from '@codemirror/state'
import { EditorView } from '@codemirror/view'
import { describe, expect, it, vi } from 'vitest'
import { mermaidCompletionSource } from './mermaidCompletion'

const declarations = `sequenceDiagram
    participant web as Web Browser
    participant blog as Blog Service
    participant account as Account Service
    participant mail as Mail Service
    participant db as Storage
`

function complete(source: string, explicit = false) {
  const pos = source.indexOf('|')
  const state = EditorState.create({ doc: source.replace('|', '') })
  return mermaidCompletionSource(new CompletionContext(state, pos, explicit))
}

describe('Mermaid entity completion', () => {
  it('inserts the participant ID without replacing the arrow or activation marker', async () => {
    const doc = `${declarations}web->>+acc`
    const parent = document.createElement('div')
    document.body.append(parent)
    const editor = new EditorView({
      parent,
      state: EditorState.create({
        doc,
        selection: { anchor: doc.length },
        extensions: [autocompletion({ override: [mermaidCompletionSource], interactionDelay: 0 })],
      }),
    })
    try {
      startCompletion(editor)
      await vi.waitFor(() => {
        expect(currentCompletions(editor.state).map((option) => option.label)).toEqual(['account'])
      })
      expect(acceptCompletion(editor)).toBe(true)
      expect(editor.state.doc.toString()).toBe(`${declarations}web->>+account`)
    } finally {
      editor.destroy()
      parent.remove()
    }
  })

  it('suggests declared IDs with aliases at the start of a message', () => {
    const result = complete(`${declarations}    ac|`)
    expect(result?.options.map((option) => option.label)).toEqual(
      expect.arrayContaining(['web', 'blog', 'account', 'mail', 'db']),
    )
    expect(result?.options.find((option) => option.label === 'account')).toMatchObject({
      detail: 'Account Service',
      type: 'variable',
    })
  })

  it.each(['->>', '->>+', '->>-', '--)', '-->>-', '--x', '<<->>'])(
    'replaces only the recipient after %s',
    (arrow) => {
      const source = `${declarations}    blog${arrow}ac|`
      const result = complete(source)
      expect(result?.from).toBe(source.indexOf('|') - 2)
      expect(result?.options.map((option) => option.label)).toContain('account')
      expect(result?.options.map((option) => option.label)).not.toContain('participant')
    },
  )

  it.each(['web->>+', 'Note over web,', 'Note right of ', 'activate ', 'deactivate '])(
    'suggests IDs at an empty reference after %s',
    (prefix) => {
      const result = complete(`${declarations}${prefix}|`)
      expect(result?.options.map((option) => option.label)).toContain('db')
    },
  )

  it('discovers implicit participants and prefers declared aliases without duplicates', () => {
    const result = complete(`sequenceDiagram
client->>api: Request
api--)worker: Notify
actor client as Customer
create participant worker as Background Worker
cl|`)
    expect(result?.options.filter((option) => option.label === 'client')).toHaveLength(1)
    expect(result?.options.find((option) => option.label === 'client')?.detail).toBe('Customer')
    expect(result?.options.find((option) => option.label === 'worker')?.detail).toBe('Background Worker')
    expect(result?.options.map((option) => option.label)).toContain('api')
  })

  it('supports IDs containing digits, underscores, and hyphens', () => {
    const source = 'sequenceDiagram\nparticipant service_2-api\nweb->>+service_2-|'
    expect(complete(source)?.from).toBe(source.indexOf('service_2-', source.indexOf('web')))
    expect(complete(source)?.options.map((option) => option.label)).toContain('service_2-api')
  })

  it.each(['%% comment ac', 'web->>db: ac', 'participant web as Web Br', 'alt Credentials fo'])(
    'does not suggest entities inside text: %s',
    (text) => expect(complete(`${declarations}${text}|`, true)).toBeNull(),
  )

  it('ignores commented declarations and refreshes suggestions after source changes', () => {
    expect(complete('sequenceDiagram\n%% participant removed\nparticipant current\ncu|')
      ?.options.map((option) => option.label)).not.toContain('removed')
    expect(complete('sequenceDiagram\nparticipant renamed\nre|')
      ?.options.map((option) => option.label)).toContain('renamed')
  })

  it('preserves diagram templates and keyword completion', () => {
    expect(complete('seq|')?.options.map((option) => option.label)).toContain('sequenceDiagram')
    expect(complete('flowchart LR\nsub|')?.options.map((option) => option.label)).toContain('subgraph')
    expect(complete('flowchart LR\n|')).toBeNull()
  })
})
