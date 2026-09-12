import {
  autocompletion,
  closeBrackets,
  closeBracketsKeymap,
  completionKeymap,
} from '@codemirror/autocomplete'
import {
  defaultKeymap,
  history,
  historyKeymap,
  indentWithTab,
} from '@codemirror/commands'
import { Compartment, EditorState } from '@codemirror/state'
import { oneDark } from '@codemirror/theme-one-dark'
import { EditorView, keymap, lineNumbers } from '@codemirror/view'
import { lintGutter, setDiagnostics } from '@codemirror/lint'
import { search, searchKeymap } from '@codemirror/search'
import { useEffect, useRef } from 'react'
import type { MermaidDiagnostic } from '../mermaid/diagnostics'
import { createMermaidDiagnostic } from './editorDiagnostics'
import { mermaidCompletionSource } from './mermaidCompletion'
import { mermaidLanguage } from './mermaidLanguage'
import './MermaidEditor.css'

interface MermaidEditorProps {
  value: string
  onChange: (value: string) => void
  diagnostic?: MermaidDiagnostic | null
  historyKey?: string
  darkMode?: boolean
}

function createEditorState(
  document: string,
  onChange: (value: string) => void,
  theme: Compartment,
  darkMode: boolean,
) {
  return EditorState.create({
    doc: document,
    extensions: [
      lineNumbers(),
      lintGutter(),
      mermaidLanguage,
      theme.of(darkMode ? oneDark : []),
      history(),
      search({ top: true }),
      closeBrackets(),
      autocompletion({ override: [mermaidCompletionSource] }),
      keymap.of([
        indentWithTab,
        ...closeBracketsKeymap,
        ...completionKeymap,
        ...defaultKeymap,
        ...historyKeymap,
        ...searchKeymap,
      ]),
      EditorView.lineWrapping,
      EditorView.contentAttributes.of({
        'aria-label': 'Mermaid diagram source',
        spellcheck: 'false',
      }),
      EditorView.updateListener.of((update) => {
        if (update.docChanged) {
          onChange(update.state.doc.toString())
        }
      }),
    ],
  })
}

export function MermaidEditor({
  value,
  onChange,
  diagnostic = null,
  historyKey = 'default',
  darkMode = false,
}: MermaidEditorProps) {
  const hostRef = useRef<HTMLDivElement>(null)
  const editorRef = useRef<EditorView>(null)
  const initialValue = useRef(value)
  const currentHistoryKey = useRef(historyKey)
  const onChangeRef = useRef(onChange)
  const themeCompartment = useRef(new Compartment())
  const darkModeRef = useRef(darkMode)

  useEffect(() => {
    onChangeRef.current = onChange
    darkModeRef.current = darkMode
  }, [onChange, darkMode])

  useEffect(() => {
    if (!hostRef.current) {
      return
    }

    const editor = new EditorView({
      parent: hostRef.current,
      state: createEditorState(
        initialValue.current,
        (nextValue) => onChangeRef.current(nextValue),
        themeCompartment.current,
        darkModeRef.current,
      ),
    })

    editorRef.current = editor
    return () => {
      editor.destroy()
      editorRef.current = null
    }
  }, [])

  useEffect(() => {
    const editor = editorRef.current
    if (
      !editor ||
      (editor.state.doc.toString() === value &&
        currentHistoryKey.current === historyKey)
    ) {
      return
    }

    // Loading another Word object starts a separate editing history.
    currentHistoryKey.current = historyKey
    editor.setState(
      createEditorState(
        value,
        (nextValue) => onChangeRef.current(nextValue),
        themeCompartment.current,
        darkModeRef.current,
      ),
    )
  }, [historyKey, value])

  useEffect(() => {
    editorRef.current?.dispatch({
      effects: themeCompartment.current.reconfigure(darkMode ? oneDark : []),
    })
  }, [darkMode])

  useEffect(() => {
    const editor = editorRef.current
    if (!editor) {
      return
    }

    const nextDiagnostic = createMermaidDiagnostic(editor.state.doc, diagnostic)
    editor.dispatch(setDiagnostics(editor.state, nextDiagnostic ? [nextDiagnostic] : []))
  }, [diagnostic, value])

  return (
    <div className="editor-shell">
      <div className="editor" ref={hostRef} />
    </div>
  )
}
