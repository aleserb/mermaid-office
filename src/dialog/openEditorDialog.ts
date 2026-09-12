import { parseDialogMessage, type ParentToDialogMessage } from './messages'
import type { DiagramSize, DiagramTheme } from '../metadata/payload'
import type { RasterizedDiagram } from '../word/insertDiagram'

export interface EditorResult {
  source: string
  theme: DiagramTheme
  size: DiagramSize
  svg: string
  raster: RasterizedDiagram
}

export function openEditorDialog(
  source: string,
  theme: DiagramTheme,
  size: DiagramSize,
  mode: 'insert' | 'update',
): Promise<EditorResult | null> {
  if (typeof Office === 'undefined' || !Office.context?.ui) {
    return Promise.reject(new Error('The expanded editor is available inside Microsoft Word.'))
  }

  const editorUrl = new URL('editor.html', window.location.href)
  editorUrl.searchParams.set('v', String(Date.now()))

  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      editorUrl.href,
      { height: 80, width: 80, displayInIframe: true },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(result.error.message))
          return
        }

        const dialog = result.value
        let settled = false
        const finish = (value: EditorResult | null) => {
          if (settled) {
            return
          }
          settled = true
          dialog.close()
          resolve(value)
        }
        const fail = (error: Error) => {
          if (settled) {
            return
          }
          settled = true
          dialog.close()
          reject(error)
        }

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (args) => {
            if (!('message' in args)) {
              return
            }
            try {
              const message = parseDialogMessage(args.message)
              if (message.type === 'ready') {
                const initialization: ParentToDialogMessage = {
                  type: 'initialize',
                  source,
                  theme,
                  size,
                  mode,
                }
                dialog.messageChild(JSON.stringify(initialization))
              } else if (message.type === 'save') {
                finish({
                  source: message.source,
                  theme: message.theme,
                  size: message.size,
                  svg: message.svg,
                  raster: message.raster,
                })
              } else {
                finish(null)
              }
            } catch (error) {
              fail(error instanceof Error ? error : new Error('Invalid editor response.'))
            }
          },
        )
        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          (args) => {
            if ('error' in args && args.error === 12006) {
              finish(null)
            } else if ('error' in args) {
              fail(new Error(`The expanded editor closed unexpectedly (${args.error}).`))
            }
          },
        )
      },
    )
  })
}
