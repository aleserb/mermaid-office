import {
  getDiagramIdFromTag,
  getDocumentSettingKey,
  parseDiagramPayload,
  type DiagramPayload,
} from '../metadata/payload'

export async function getSelectedDiagram(): Promise<DiagramPayload | null> {
  if (typeof Word === 'undefined') {
    return null
  }

  return Word.run(async (context) => {
    const selection = context.document.getSelection()
    const directParent = selection.parentContentControlOrNullObject
    const selectedPicture = selection.inlinePictures.getFirstOrNullObject()
    directParent.load('tag')
    await context.sync()

    let contentControl = directParent
    if (directParent.isNullObject) {
      await context.sync()
      if (selectedPicture.isNullObject) {
        return null
      }

      contentControl = selectedPicture.parentContentControlOrNullObject
      contentControl.load('tag')
      await context.sync()
    }

    if (contentControl.isNullObject) {
      return null
    }

    const id = getDiagramIdFromTag(contentControl.tag)
    if (!id) {
      return null
    }
    const setting = context.document.settings.getItemOrNullObject(getDocumentSettingKey(id))
    setting.load('value')
    await context.sync()

    return setting.isNullObject ? null : parseDiagramPayload(String(setting.value))
  })
}

export function watchSelectedDiagram(
  onSelected: (payload: DiagramPayload | null) => void,
  onError: (error: Error) => void,
): () => void {
  if (typeof Office === 'undefined' || !Office.context?.document) {
    return () => undefined
  }

  let active = true
  let checking = false
  let queued = false

  const checkSelection = async () => {
    if (checking) {
      queued = true
      return
    }

    checking = true
    try {
      const payload = await getSelectedDiagram()
      if (active) {
        onSelected(payload)
      }
    } catch (error) {
      if (active) {
        onError(error instanceof Error ? error : new Error('Unable to read selected diagram.'))
      }
    } finally {
      checking = false
      if (active && queued) {
        queued = false
        void checkSelection()
      }
    }
  }

  const handler = () => {
    void checkSelection()
  }

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    handler,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed && active) {
        onError(new Error(result.error.message))
      }
    },
  )
  void checkSelection()

  return () => {
    active = false
    Office.context.document.removeHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      { handler },
    )
  }
}
