import { EditorState } from '@codemirror/state'
import { EditorView, keymap, lineNumbers } from '@codemirror/view'
import { indentWithTab } from '@codemirror/commands'
import { useEffect, useEffectEvent, useRef } from 'react'

interface MermaidEditorProps {
  value: string
  onChange: (value: string) => void
}

export function MermaidEditor({ value, onChange }: MermaidEditorProps) {
  const hostRef = useRef<HTMLDivElement>(null)
  const editorRef = useRef<EditorView>(null)
  const initialValue = useRef(value)
  const handleChange = useEffectEvent((nextValue: string) => {
    onChange(nextValue)
  })

  useEffect(() => {
    if (!hostRef.current) {
      return
    }

    const editor = new EditorView({
      parent: hostRef.current,
      state: EditorState.create({
        doc: initialValue.current,
        extensions: [
          lineNumbers(),
          keymap.of([indentWithTab]),
          EditorView.lineWrapping,
          EditorView.contentAttributes.of({
            'aria-label': 'Mermaid diagram source',
            spellcheck: 'false',
          }),
          EditorView.updateListener.of((update) => {
            if (update.docChanged) {
              handleChange(update.state.doc.toString())
            }
          }),
        ],
      }),
    })

    editorRef.current = editor
    return () => {
      editor.destroy()
      editorRef.current = null
    }
  }, [])

  useEffect(() => {
    const editor = editorRef.current
    if (!editor || editor.state.doc.toString() === value) {
      return
    }

    editor.dispatch({
      changes: { from: 0, to: editor.state.doc.length, insert: value },
    })
  }, [value])

  return <div className="editor" ref={hostRef} />
}
