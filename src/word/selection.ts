import {
  getDiagramIdFromTag,
  getContentControlTag,
  getDocumentSettingKey,
  parseDiagramPayload,
  type DiagramPayload,
} from '../metadata/payload'
import { readPayloadFromImage } from '../metadata/imageMetadata'

function copyWithNewId(payload: DiagramPayload): DiagramPayload {
  return { ...payload, id: crypto.randomUUID() }
}

function configureRecoveredContentControl(
  contentControl: Word.ContentControl,
  payload: DiagramPayload,
) {
  contentControl.tag = getContentControlTag(payload.id)
  contentControl.title = 'Mermaid diagram'
  contentControl.appearance = Word.ContentControlAppearance.hidden
  contentControl.cannotDelete = false
  contentControl.cannotEdit = false
  contentControl.select()
}

async function readPicturePayload(
  context: Word.RequestContext,
  picture: Word.InlinePicture,
): Promise<DiagramPayload | null> {
  const image = picture.getBase64ImageSrc()
  await context.sync()
  return readPayloadFromImage(image.value)
}

export async function getSelectedDiagram(): Promise<DiagramPayload | null> {
  if (typeof Word === 'undefined') {
    return null
  }

  return Word.run(async (context) => {
    const selection = context.document.getSelection()
    const selectedPicture = selection.inlinePictures.getFirstOrNullObject()
    await context.sync()

    // Enter can extend a diagram's hidden control to include blank paragraphs.
    // Being inside that control is not the same as selecting its picture.
    if (selectedPicture.isNullObject) {
      return null
    }

    let contentControl = selectedPicture.parentContentControlOrNullObject
    contentControl.load('tag')
    await context.sync()

    if (contentControl.isNullObject) {
      const embeddedPayload = await readPicturePayload(context, selectedPicture)
      if (!embeddedPayload) {
        return null
      }

      const recoveredPayload = copyWithNewId(embeddedPayload)
      contentControl = selectedPicture.insertContentControl()
      configureRecoveredContentControl(contentControl, recoveredPayload)
      context.document.settings.add(
        getDocumentSettingKey(recoveredPayload.id),
        JSON.stringify(recoveredPayload),
      )
      await context.sync()
      return recoveredPayload
    }

    const id = getDiagramIdFromTag(contentControl.tag)
    if (!id) {
      return null
    }

    const setting = context.document.settings.getItemOrNullObject(getDocumentSettingKey(id))
    const matchingControls = context.document.contentControls.getByTag(
      getContentControlTag(id),
    )
    setting.load('value')
    matchingControls.load('items')
    await context.sync()

    let storedPayload: DiagramPayload | null = null
    let storedPayloadError: unknown
    if (!setting.isNullObject) {
      try {
        storedPayload = parseDiagramPayload(String(setting.value))
      } catch (error) {
        storedPayloadError = error
      }
    }

    const needsRecovery = !storedPayload || matchingControls.items.length > 1
    if (!needsRecovery) {
      return storedPayload
    }

    const picture = contentControl.inlinePictures.getFirstOrNullObject()
    await context.sync()
    const embeddedPayload = picture.isNullObject
      ? null
      : await readPicturePayload(context, picture)
    const recoveredPayload = embeddedPayload ?? storedPayload
    if (!recoveredPayload) {
      if (storedPayloadError instanceof Error) {
        throw storedPayloadError
      }
      return null
    }

    const uniquePayload =
      matchingControls.items.length > 1
        ? copyWithNewId(recoveredPayload)
        : { ...recoveredPayload, id }
    configureRecoveredContentControl(contentControl, uniquePayload)
    context.document.settings.add(
      getDocumentSettingKey(uniquePayload.id),
      JSON.stringify(uniquePayload),
    )
    await context.sync()
    return uniquePayload
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
