import { embedPayloadInPng } from '../metadata/pngMetadata'
import {
  createDiagramPayload,
  getContentControlTag,
  getDocumentSettingKey,
  type DiagramFormat,
  type DiagramPayload,
  type DiagramTheme,
} from '../metadata/payload'
import { embedPayloadInSvg } from '../metadata/svgMetadata'

export type { DiagramFormat } from '../metadata/payload'

interface RasterizedDiagram {
  base64: string
  width: number
  height: number
}

export function isSvgInsertionSupported(): boolean {
  return (
    typeof Office !== 'undefined' &&
    Office.context.requirements.isSetSupported('ImageCoercion', '1.2')
  )
}

function setSelectedData(data: string, coercionType: Office.CoercionType): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(data, { coercionType }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve()
      } else {
        reject(new Error(result.error.message))
      }
    })
  })
}

export async function rasterizeSvg(svg: string): Promise<RasterizedDiagram> {
  const svgDocument = new DOMParser().parseFromString(svg, 'image/svg+xml')
  const root = svgDocument.documentElement
  const viewBox = root.getAttribute('viewBox')?.split(/\s+/).map(Number)
  const sourceWidth = viewBox?.[2] || Number.parseFloat(root.getAttribute('width') || '') || 1200
  const sourceHeight = viewBox?.[3] || Number.parseFloat(root.getAttribute('height') || '') || 800
  const scale = Math.min(2, 4096 / Math.max(sourceWidth, sourceHeight))
  const width = Math.max(1, Math.round(sourceWidth * scale))
  const height = Math.max(1, Math.round(sourceHeight * scale))
  root.setAttribute('width', String(sourceWidth))
  root.setAttribute('height', String(sourceHeight))
  root.style.removeProperty('max-width')
  const normalizedSvg = new XMLSerializer().serializeToString(root)
  const url = URL.createObjectURL(new Blob([normalizedSvg], { type: 'image/svg+xml' }))

  try {
    const image = new Image()
    image.src = url
    await image.decode()

    const canvas = window.document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    const context = canvas.getContext('2d')
    if (!context) {
      throw new Error('This browser cannot create the PNG fallback.')
    }

    context.scale(scale, scale)
    context.drawImage(image, 0, 0, sourceWidth, sourceHeight)
    return {
      base64: canvas.toDataURL('image/png').split(',', 2)[1],
      width: sourceWidth,
      height: sourceHeight,
    }
  } finally {
    URL.revokeObjectURL(url)
  }
}

export async function svgToPngBase64(svg: string): Promise<string> {
  return (await rasterizeSvg(svg)).base64
}

function configureDiagramContentControl(
  picture: Word.InlinePicture,
  payload: DiagramPayload,
): Word.ContentControl {
  picture.altTextTitle = 'Mermaid diagram'
  picture.altTextDescription = 'Diagram created with Mermaid Office.'

  const contentControl = picture.insertContentControl()
  contentControl.tag = getContentControlTag(payload.id)
  contentControl.title = 'Mermaid diagram'
  contentControl.appearance = Word.ContentControlAppearance.hidden
  contentControl.cannotDelete = false
  contentControl.cannotEdit = false
  contentControl.select()
  return contentControl
}

async function insertPngObject(
  raster: RasterizedDiagram,
  payload: DiagramPayload,
): Promise<void> {
  await Word.run(async (context) => {
    const selection = context.document.getSelection()
    const picture = selection.insertInlinePictureFromBase64(
      raster.base64,
      Word.InsertLocation.replace,
    )
    const dimensions = fitDiagram(raster.width, raster.height)
    picture.width = dimensions.width
    picture.height = dimensions.height
    configureDiagramContentControl(picture, payload)
    context.document.settings.add(getDocumentSettingKey(payload.id), JSON.stringify(payload))
    await context.sync()
  })
}

export async function updateDiagram(
  svg: string,
  existing: DiagramPayload,
  source: string,
  theme: DiagramTheme = existing.theme,
): Promise<DiagramFormat> {
  if (typeof Word === 'undefined') {
    throw new Error('Open Mermaid Office inside Microsoft Word to update a diagram.')
  }

  const payload: DiagramPayload = {
    ...existing,
    source,
    theme,
    format: 'png',
  }
  const raster = await rasterizeSvg(svg)
  const png = embedPayloadInPng(raster.base64, payload)

  await Word.run(async (context) => {
    const selection = context.document.getSelection()
    const directParent = selection.parentContentControlOrNullObject
    const selectedPicture = selection.inlinePictures.getFirstOrNullObject()
    directParent.load('tag')
    await context.sync()

    let contentControl = directParent
    if (directParent.isNullObject) {
      if (selectedPicture.isNullObject) {
        throw new Error('Select the Mermaid diagram you want to update.')
      }
      contentControl = selectedPicture.parentContentControlOrNullObject
      contentControl.load('tag')
      await context.sync()
    }

    if (
      contentControl.isNullObject ||
      contentControl.tag !== getContentControlTag(existing.id)
    ) {
      throw new Error('Select the same Mermaid diagram before updating it.')
    }

    const existingPicture = contentControl.inlinePictures.getFirstOrNullObject()
    existingPicture.load('width,altTextTitle,altTextDescription')
    await context.sync()

    if (existingPicture.isNullObject) {
      throw new Error('The selected Mermaid diagram no longer contains a picture.')
    }

    const replacement = contentControl.insertInlinePictureFromBase64(
      png,
      Word.InsertLocation.replace,
    )
    const aspectRatio = raster.height / raster.width
    replacement.width = existingPicture.width
    replacement.height = existingPicture.width * aspectRatio
    replacement.altTextTitle = existingPicture.altTextTitle || 'Mermaid diagram'
    replacement.altTextDescription =
      existingPicture.altTextDescription || 'Diagram created with Mermaid Office.'
    context.document.settings.add(
      getDocumentSettingKey(payload.id),
      JSON.stringify(payload),
    )
    contentControl.select()
    await context.sync()
  })

  return 'png'
}

export function fitDiagram(
  width: number,
  height: number,
  maxWidth = 500,
  maxHeight = 650,
): { width: number; height: number } {
  const scale = Math.min(1, maxWidth / width, maxHeight / height)
  return {
    width: width * scale,
    height: height * scale,
  }
}

async function wrapSelectedSvgObject(payload: DiagramPayload): Promise<void> {
  await Word.run(async (context) => {
    const picture = context.document.getSelection().inlinePictures.getFirstOrNullObject()
    await context.sync()

    if (picture.isNullObject) {
      throw new Error('Word inserted the SVG but did not expose it as the current picture.')
    }

    const contentControl = configureDiagramContentControl(picture, payload)
    const setting = context.document.settings.add(
      getDocumentSettingKey(payload.id),
      JSON.stringify(payload),
    )

    try {
      await context.sync()
    } catch (insertError) {
      contentControl.delete(false)
      setting.delete()
      try {
        await context.sync()
      } catch (rollbackError) {
        throw new AggregateError(
          [insertError, rollbackError],
          'SVG metadata failed and Word could not fully roll back the inserted diagram.',
        )
      }
      throw insertError
    }
  })
}

export async function insertDiagram(
  svg: string,
  source: string,
  theme: DiagramTheme = 'default',
): Promise<DiagramFormat> {
  if (typeof Office === 'undefined' || !Office.context?.document) {
    throw new Error('Open Mermaid Office inside Microsoft Word to insert a diagram.')
  }

  if (isSvgInsertionSupported()) {
    const payload = createDiagramPayload(source, 'svg', theme)
    await setSelectedData(embedPayloadInSvg(svg, payload), Office.CoercionType.XmlSvg)
    await wrapSelectedSvgObject(payload)
    return 'svg'
  }

  const payload = createDiagramPayload(source, 'png', theme)
  const raster = await rasterizeSvg(svg)
  const png = embedPayloadInPng(raster.base64, payload)
  await insertPngObject({ ...raster, base64: png }, payload)
  return 'png'
}
