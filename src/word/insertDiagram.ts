import { embedPayloadInPng, setPngPhysicalWidth } from '../metadata/pngMetadata'
import { configureDiagramContentControl, suppressDiagramPlaceholder } from './contentControls'
import { createDiagramPictureOoxml } from './pictureOoxml'
import {
  createDiagramPayload,
  getContentControlTag,
  getDiagramIdFromTag,
  getDocumentSettingKey,
  parseDiagramPayload,
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

export interface DiagramInsertionOptions {
  requireEmptySelection?: boolean
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

const MAX_RASTER_DIMENSION = 8192
const MAX_RASTER_PIXELS = 32 * 1024 * 1024

export function getRasterDimensions(
  width: number,
  height: number,
  size?: DiagramSize,
): { width: number; height: number } {
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    throw new Error('Diagram dimensions must be positive finite numbers.')
  }
  // Eight pixels per displayed CSS pixel cover 400% zoom on a 2x-density screen.
  // Dense diagrams also target 3x native detail, within the same resource limits.
  const preferredScale = Math.max(3, size ? DIAGRAM_WIDTHS[size] * (96 / 72) * 8 / width : 3)
  const scale = Math.min(
    preferredScale,
    MAX_RASTER_DIMENSION / width,
    MAX_RASTER_DIMENSION / height,
    Math.sqrt(MAX_RASTER_PIXELS / width / height),
  )
  return {
    width: Math.max(1, Math.floor(width * scale)),
    height: Math.max(1, Math.floor(height * scale)),
  }
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
    const logicalWidth = croppedWidth / scale
    const logicalHeight = croppedHeight / scale
    const dimensions = getRasterDimensions(logicalWidth, logicalHeight, size)
    canvas.width = 1
    canvas.height = 1
    const output = window.document.createElement('canvas')
    output.width = dimensions.width
    output.height = dimensions.height
    try {
      const outputContext = output.getContext('2d')
      if (!outputContext) {
        throw new Error('This browser cannot crop the PNG fallback.')
      }
      // Rasterize the vector source at the final resolution, not the bounds-probe
      // bitmap: enlarging that bitmap would add pixels without adding detail.
      outputContext.drawImage(
        image,
        left / scale,
        top / scale,
        logicalWidth,
        logicalHeight,
        0,
        0,
        output.width,
        output.height,
      )

      const dataUrl = output.toDataURL('image/png')
      const base64 = dataUrl.split(',', 2)[1]
      if (!dataUrl.startsWith('data:image/png;base64,') || !base64) {
        throw new Error('This browser cannot export the diagram at this PNG resolution.')
      }
      return {
        base64,
        width: logicalWidth,
        height: logicalHeight,
      }
    } finally {
      // Release the large backing store immediately between live updates.
      output.width = 1
      output.height = 1
    }
  } finally {
    URL.revokeObjectURL(url)
  }
}

export async function svgToPngBase64(svg: string): Promise<string> {
  return (await rasterizeSvg(svg)).base64
}

async function separateEnclosingDiagram(
  context: Word.RequestContext,
  selection: Word.Range = context.document.getSelection(),
): Promise<void> {
  const enclosing = selection.parentContentControlOrNullObject
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
  options: DiagramInsertionOptions = {},
): Promise<void> {
  await Word.run(async (context) => {
    const insertionSelection = options.requireEmptySelection
      ? context.document.getSelection()
      : undefined
    if (insertionSelection) {
      await requireInsertionCursor(context, insertionSelection)
    }
    await separateEnclosingDiagram(context, insertionSelection)
    // Unwrapping an enclosing diagram can invalidate the previous selection path.
    const selection = context.document.getSelection()
    if (options.requireEmptySelection) {
      await requireInsertionCursor(context, selection)
    }
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
    contentControl.insertOoxml(
      createDiagramPictureOoxml(png, dimensions),
      Word.InsertLocation.replace,
    )
    context.document.settings.add(getDocumentSettingKey(payload.id), JSON.stringify(payload))
    await context.sync()
  })
}

async function requireInsertionCursor(context: Word.RequestContext, selection: Word.Range): Promise<void> {
  const picture = selection.inlinePictures.getFirstOrNullObject()
  selection.load('isEmpty')
  await context.sync()
  if (!selection.isEmpty || !picture.isNullObject) {
    throw new Error('Place the cursor without selecting text or a picture before inserting a diagram.')
  }
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

    suppressDiagramPlaceholder(contentControl)
    const replacement = replaceDiagramPicture(context, existingPicture, raster, payload, applySize, {
      altTextTitle: existingPicture.altTextTitle || 'Mermaid diagram',
      altTextDescription: existingPicture.altTextDescription || 'Diagram created with Mermaid Office.',
    })
    replacement.select()
    await context.sync()
  })

  return 'png'
}

function replaceDiagramPicture(
  context: Word.RequestContext,
  existingPicture: Word.InlinePicture,
  raster: RasterizedDiagram,
  payload: DiagramPayload,
  applySize: boolean,
  altText: Pick<Word.InlinePicture, 'altTextTitle' | 'altTextDescription'> = existingPicture,
): Word.Range {
  const { png, ...properties } = prepareDiagramPicture(existingPicture, raster, payload, applySize, altText)
  // Replace only the picture: deleting its control can invalidate Word's range
  // and remove surrounding text or other diagrams.
  const replacement = existingPicture.getRange().insertOoxml(
    createDiagramPictureOoxml(png, properties),
    Word.InsertLocation.replace,
  )
  context.document.settings.add(getDocumentSettingKey(payload.id), JSON.stringify(payload))
  return replacement
}

function prepareDiagramPicture(
  existingPicture: Word.InlinePicture,
  raster: RasterizedDiagram,
  payload: DiagramPayload,
  applySize: boolean,
  altText: Pick<Word.InlinePicture, 'altTextTitle' | 'altTextDescription'> = existingPicture,
) {
  const dimensions = applySize
    ? fitDiagram(raster.width, raster.height, DIAGRAM_WIDTHS[payload.size], 650, true)
    : {
        width: existingPicture.width,
        height: existingPicture.width * (raster.height / raster.width),
      }
  const png = setPngPhysicalWidth(
    embedPayloadInPng(raster.base64, payload),
    dimensions.width,
  )
  return {
    png,
    ...dimensions,
    altTextTitle: altText.altTextTitle,
    altTextDescription: altText.altTextDescription,
  }
}

async function getDiagramUpdateTarget(
  context: Word.RequestContext,
  existing: DiagramPayload,
  expectedControlId?: number,
) {
  const controls = context.document.contentControls.getByTag(getContentControlTag(existing.id))
  controls.load('items')
  await context.sync()
  if (controls.items.length === 0) {
    throw new Error('The Mermaid diagram was deleted or can no longer be found.')
  }
  if (controls.items.length !== 1) {
    throw new Error('Multiple Mermaid diagrams share this ID. Select the diagram again to edit it.')
  }

  const control = controls.items[0]
  const pictures = control.inlinePictures
  const setting = context.document.settings.getItemOrNullObject(getDocumentSettingKey(existing.id))
  control.load('id')
  pictures.load('items')
  setting.load('value')
  await context.sync()
  if (expectedControlId !== undefined && control.id !== expectedControlId) {
    throw new Error('The Mermaid diagram was replaced while updating. Select it again to edit it.')
  }
  if (pictures.items.length !== 1) {
    throw new Error('The Mermaid diagram must contain exactly one picture to update it safely.')
  }
  if (!setting.isNullObject) {
    const stored = parseDiagramPayload(String(setting.value))
    if (
      stored.id !== existing.id ||
      stored.source !== existing.source ||
      stored.theme !== existing.theme ||
      stored.size !== existing.size ||
      stored.format !== existing.format
    ) {
      throw new Error('The Mermaid diagram changed in another editor. Select it again before updating.')
    }
  }
  return { control, picture: pictures.items[0] }
}

export async function updateDiagramById(
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

  const payload: DiagramPayload = { ...existing, source, theme, size, format: 'png' }
  const raster = renderedRaster ?? (await rasterizeSvg(svg, size))
  const update = await Word.run(async (context) => {
    const { control, picture } = await getDiagramUpdateTarget(context, existing)
    picture.load('altTextTitle,altTextDescription,width')
    await context.sync()
    const { png, ...properties } = prepareDiagramPicture(picture, raster, payload, applySize)
    suppressDiagramPlaceholder(control)
    picture.getRange().insertInlinePictureFromBase64(png, Word.InsertLocation.replace)
    await context.sync()
    return { controlId: control.id, properties }
  })

  // Insertion results and collection items can retain old bitmap geometry.
  // Resolve the picture directly by control ID in a new request before sizing.
  // This avoids OOXML imports and their blocking dialog during live edits.
  await Word.run(async (context) => {
    await getDiagramUpdateTarget(context, existing, update.controlId)
    const picture = context.document.contentControls.getById(update.controlId).inlinePictures.getFirst()
    picture.load('width,height')
    await context.sync()
    picture.lockAspectRatio = false
    picture.width = update.properties.width
    picture.height = update.properties.height
    picture.lockAspectRatio = true
    picture.altTextTitle = update.properties.altTextTitle
    picture.altTextDescription = update.properties.altTextDescription
    context.document.settings.add(getDocumentSettingKey(payload.id), JSON.stringify(payload))
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
  return (await insertDiagramWithPayload(svg, source, theme, size, renderedRaster)).format
}

export async function insertDiagramWithPayload(
  svg: string,
  source: string,
  theme: DiagramTheme = 'default',
  size: DiagramSize = 'medium',
  renderedRaster?: RasterizedDiagram,
  options: DiagramInsertionOptions = {},
): Promise<DiagramPayload> {
  if (typeof Office === 'undefined' || !Office.context?.document) {
    throw new Error('Open Mermaid Office inside Microsoft Word to insert a diagram.')
  }

  const payload = createDiagramPayload(source, 'png', theme, size)
  const raster = renderedRaster ?? (await rasterizeSvg(svg, size))
  const png = embedPayloadInPng(raster.base64, payload)
  await insertPngObject({ ...raster, base64: png }, payload, options)
  return payload
}
