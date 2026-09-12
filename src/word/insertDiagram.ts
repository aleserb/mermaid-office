import { embedPayloadInPng, setPngPhysicalWidth } from '../metadata/pngMetadata'
import {
  createDiagramPayload,
  getContentControlTag,
  getDiagramIdFromTag,
  getDocumentSettingKey,
  type DiagramFormat,
  type DiagramPayload,
  type DiagramSize,
  type DiagramTheme,
} from '../metadata/payload'

export type { DiagramFormat } from '../metadata/payload'

export interface RasterizedDiagram {
  base64: string
  width: number
  height: number
}

interface PixelBounds {
  left: number
  top: number
  width: number
  height: number
}

const DIAGRAM_WIDTHS: Record<DiagramSize, number> = {
  small: 216,
  medium: 324,
  large: 396,
  'page-width': 468,
}

export function findVisiblePixelBounds(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
): PixelBounds | null {
  let left = width
  let top = height
  let right = -1
  let bottom = -1

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      if (pixels[(y * width + x) * 4 + 3] === 0) {
        continue
      }
      left = Math.min(left, x)
      top = Math.min(top, y)
      right = Math.max(right, x)
      bottom = Math.max(bottom, y)
    }
  }

  return right < left
    ? null
    : {
        left,
        top,
        width: right - left + 1,
        height: bottom - top + 1,
      }
}

export function normalizeSvgDimensions(svg: string): {
  svg: string
  width: number
  height: number
} {
  const svgDocument = new DOMParser().parseFromString(svg, 'image/svg+xml')
  const root = svgDocument.documentElement
  const viewBox = root.getAttribute('viewBox')?.split(/\s+/).map(Number)
  const sourceWidth = viewBox?.[2] || Number.parseFloat(root.getAttribute('width') || '') || 1200
  const sourceHeight = viewBox?.[3] || Number.parseFloat(root.getAttribute('height') || '') || 800
  root.setAttribute('width', String(sourceWidth))
  root.setAttribute('height', String(sourceHeight))
  root.style.removeProperty('max-width')

  return {
    svg: new XMLSerializer().serializeToString(root),
    width: sourceWidth,
    height: sourceHeight,
  }
}

export async function rasterizeSvg(
  svg: string,
  size?: DiagramSize,
): Promise<RasterizedDiagram> {
  const normalized = normalizeSvgDimensions(svg)
  const scale = Math.min(2, 4096 / Math.max(normalized.width, normalized.height))
  const width = Math.max(1, Math.round(normalized.width * scale))
  const height = Math.max(1, Math.round(normalized.height * scale))
  const url = URL.createObjectURL(new Blob([normalized.svg], { type: 'image/svg+xml' }))

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
    context.drawImage(image, 0, 0, normalized.width, normalized.height)

    const bounds = findVisiblePixelBounds(
      context.getImageData(0, 0, width, height).data,
      width,
      height,
    )
    if (!bounds) {
      throw new Error('The rendered diagram does not contain any visible content.')
    }

    const padding = Math.max(1, Math.round(scale * 4))
    const left = Math.max(0, bounds.left - padding)
    const top = Math.max(0, bounds.top - padding)
    const right = Math.min(width, bounds.left + bounds.width + padding)
    const bottom = Math.min(height, bounds.top + bounds.height + padding)
    const croppedWidth = right - left
    const croppedHeight = bottom - top
    const targetWidth = size
      ? Math.round(DIAGRAM_WIDTHS[size] * (96 / 72) * 2)
      : croppedWidth
    const outputScale = Math.min(
      targetWidth / croppedWidth,
      4096 / croppedWidth,
      4096 / croppedHeight,
    )
    const output = window.document.createElement('canvas')
    output.width = Math.max(1, Math.round(croppedWidth * outputScale))
    output.height = Math.max(1, Math.round(croppedHeight * outputScale))
    const outputContext = output.getContext('2d')
    if (!outputContext) {
      throw new Error('This browser cannot crop the PNG fallback.')
    }
    outputContext.drawImage(
      canvas,
      left,
      top,
      croppedWidth,
      croppedHeight,
      0,
      0,
      output.width,
      output.height,
    )

    return {
      base64: output.toDataURL('image/png').split(',', 2)[1],
      width: croppedWidth / scale,
      height: croppedHeight / scale,
    }
  } finally {
    URL.revokeObjectURL(url)
  }
}

export async function svgToPngBase64(svg: string): Promise<string> {
  return (await rasterizeSvg(svg)).base64
}

function configureDiagramPicture(picture: Word.InlinePicture) {
  picture.altTextTitle = 'Mermaid diagram'
  picture.altTextDescription = 'Diagram created with Mermaid Office.'
}

function configureDiagramContentControl(
  contentControl: Word.ContentControl,
  payload: Pick<DiagramPayload, 'id'>,
) {
  contentControl.tag = getContentControlTag(payload.id)
  contentControl.title = 'Mermaid diagram'
  contentControl.appearance = Word.ContentControlAppearance.hidden
  contentControl.cannotDelete = false
  contentControl.cannotEdit = false
}

async function separateEnclosingDiagram(context: Word.RequestContext): Promise<void> {
  const enclosing = context.document.getSelection().parentContentControlOrNullObject
  enclosing.load('tag')
  await context.sync()
  const id = enclosing.isNullObject ? null : getDiagramIdFromTag(enclosing.tag)
  if (!id) {
    return
  }

  const picture = enclosing.inlinePictures.getFirstOrNullObject()
  await context.sync()
  if (picture.isNullObject) {
    enclosing.delete(true)
    await context.sync()
    return
  }

  // Keep typed paragraphs outside the old diagram so a new diagram is not nested
  // inside it (and subsequently deleted when the old diagram is updated).
  const pictureRange = picture.getRange()
  pictureRange.track()
  await context.sync()
  enclosing.delete(true)
  await context.sync()
  const isolated = pictureRange.insertContentControl()
  configureDiagramContentControl(isolated, { id })
  await context.sync()
  pictureRange.untrack()
  await context.sync()
}

export async function insertPngObject(
  raster: RasterizedDiagram,
  payload: DiagramPayload,
): Promise<void> {
  await Word.run(async (context) => {
    await separateEnclosingDiagram(context)
    const selection = context.document.getSelection()
    const contentControl = selection.insertContentControl()
    configureDiagramContentControl(contentControl, payload)
    await context.sync()

    const dimensions = fitDiagram(
      raster.width,
      raster.height,
      DIAGRAM_WIDTHS[payload.size],
      650,
      true,
    )
    const png = setPngPhysicalWidth(raster.base64, dimensions.width)
    const picture = contentControl.insertInlinePictureFromBase64(
      png,
      Word.InsertLocation.replace,
    )
    picture.width = dimensions.width
    picture.height = dimensions.height
    configureDiagramPicture(picture)
    await context.sync()

    // Word can retain the placeholder frame dimensions for the first image placed in a
    // new content control. Replacing it after the control contains a picture forces the
    // same layout pass used when updating an existing diagram.
    const replacement = contentControl.insertInlinePictureFromBase64(
      png,
      Word.InsertLocation.replace,
    )
    replacement.width = dimensions.width
    replacement.height = dimensions.height
    configureDiagramPicture(replacement)
    context.document.settings.add(getDocumentSettingKey(payload.id), JSON.stringify(payload))
    await context.sync()
  })
}

export async function updateDiagram(
  svg: string,
  existing: DiagramPayload,
  source: string,
  theme: DiagramTheme = existing.theme,
  size: DiagramSize = existing.size,
  applySize = false,
  renderedRaster?: RasterizedDiagram,
): Promise<DiagramFormat> {
  if (typeof Word === 'undefined') {
    throw new Error('Open Mermaid Office inside Microsoft Word to update a diagram.')
  }

  const payload: DiagramPayload = {
    ...existing,
    source,
    theme,
    size,
    format: 'png',
  }
  const raster = renderedRaster ?? (await rasterizeSvg(svg, size))

  await Word.run(async (context) => {
    const selection = context.document.getSelection()
    const directParent = selection.parentContentControlOrNullObject
    const selectedPicture = selection.inlinePictures.getFirstOrNullObject()
    directParent.load('tag')
    await context.sync()

    if (selectedPicture.isNullObject) {
      throw new Error('Select the Mermaid diagram you want to update.')
    }

    let contentControl = directParent
    if (directParent.isNullObject) {
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

    const existingPicture = selectedPicture
    existingPicture.load('altTextTitle,altTextDescription,width')
    await context.sync()

    if (existingPicture.isNullObject) {
      throw new Error('The selected Mermaid diagram no longer contains a picture.')
    }

    const dimensions = applySize
      ? fitDiagram(
          raster.width,
          raster.height,
          DIAGRAM_WIDTHS[size],
          650,
          true,
        )
      : {
          width: existingPicture.width,
          height: existingPicture.width * (raster.height / raster.width),
        }
    const png = setPngPhysicalWidth(
      embedPayloadInPng(raster.base64, payload),
      dimensions.width,
    )
    // Replace only the selected picture. A hidden control may also contain text
    // or other diagrams, and deleting it can invalidate Word's insertion range.
    const replacement = existingPicture.getRange().insertInlinePictureFromBase64(
      png,
      Word.InsertLocation.replace,
    )
    replacement.width = dimensions.width
    replacement.height = dimensions.height
    replacement.altTextTitle = existingPicture.altTextTitle || 'Mermaid diagram'
    replacement.altTextDescription =
      existingPicture.altTextDescription || 'Diagram created with Mermaid Office.'
    context.document.settings.add(
      getDocumentSettingKey(payload.id),
      JSON.stringify(payload),
    )
    replacement.getRange().select()
    await context.sync()
  })

  return 'png'
}

export function fitDiagram(
  width: number,
  height: number,
  maxWidth = 500,
  maxHeight = 650,
  allowUpscale = false,
): { width: number; height: number } {
  const scale = Math.min(
    allowUpscale ? Number.POSITIVE_INFINITY : 1,
    maxWidth / width,
    maxHeight / height,
  )
  return {
    width: width * scale,
    height: height * scale,
  }
}

export async function insertDiagram(
  svg: string,
  source: string,
  theme: DiagramTheme = 'default',
  size: DiagramSize = 'medium',
  renderedRaster?: RasterizedDiagram,
): Promise<DiagramFormat> {
  if (typeof Office === 'undefined' || !Office.context?.document) {
    throw new Error('Open Mermaid Office inside Microsoft Word to insert a diagram.')
  }

  const payload = createDiagramPayload(source, 'png', theme, size)
  const raster = renderedRaster ?? (await rasterizeSvg(svg, size))
  const png = embedPayloadInPng(raster.base64, payload)
  await insertPngObject({ ...raster, base64: png }, payload)
  return 'png'
}
