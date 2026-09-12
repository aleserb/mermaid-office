import { embedPayloadInPng, getPngDimensions, setPngPhysicalWidth } from '../metadata/pngMetadata'
import { getSvgDimensions } from '../mermaid/svgDimensions'
import { configureDiagramContentControl, suppressDiagramPlaceholder } from './contentControls'
import { createDiagramPictureOoxml, type PictureOptions } from './pictureOoxml'
import { getDiagramSettings, sameDiagramSettings, type DiagramSettings, type ImageQuality } from '../metadata/diagramSettings'
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

const RASTER_LIMITS = {
  auto: { dimension: 4096, pixels: 4 * 1024 * 1024 },
  standard: { dimension: 4096, pixels: 4 * 1024 * 1024 },
  high: { dimension: 8192, pixels: 32 * 1024 * 1024 },
} satisfies Record<ImageQuality, { dimension: number; pixels: number }>

export function getRasterDimensions(
  width: number,
  height: number,
  sizeOrWidth?: DiagramSize | number,
  quality: ImageQuality = 'auto',
): { width: number; height: number } {
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    throw new Error('Diagram dimensions must be positive finite numbers.')
  }
  if (typeof sizeOrWidth === 'number' && (!Number.isFinite(sizeOrWidth) || sizeOrWidth <= 0)) {
    throw new Error('Diagram display width must be a positive finite number.')
  }
  const displayWidth = typeof sizeOrWidth === 'number'
    ? sizeOrWidth
    : sizeOrWidth ? fitDiagram(width, height, DIAGRAM_WIDTHS[sizeOrWidth], 650, true).width : 0
  // Word widths are in points. Standard targets a 2x-density display; Auto also
  // retains native SVG detail, while High reserves detail for extreme zoom.
  const displayScale = displayWidth > 0 ? displayWidth * (96 / 72) * 2 / width : 1
  const preferredScale = quality === 'high'
    ? Math.max(3, displayWidth > 0 ? displayScale * 4 : 3)
    : quality === 'standard' ? displayScale : Math.max(1, displayScale)
  const limits = RASTER_LIMITS[quality]
  const scale = Math.min(
    preferredScale,
    limits.dimension / width,
    limits.dimension / height,
    Math.sqrt(limits.pixels / width / height),
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
  const { width: sourceWidth, height: sourceHeight } = getSvgDimensions(root)
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
  sizeOrWidth?: DiagramSize | number,
  quality: ImageQuality = 'auto',
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
    const dimensions = getRasterDimensions(logicalWidth, logicalHeight, sizeOrWidth, quality)
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

export async function svgToPngBase64(svg: string, quality: ImageQuality = 'auto'): Promise<string> {
  return (await rasterizeSvg(svg, undefined, quality)).base64
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
    await insertDiagramPicture(context, contentControl, png, dimensions, true)
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
  settings?: DiagramSettings,
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
    ...(settings ? { settings: { ...settings } } : {}),
  }
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

    const raster = renderedRaster ?? (await rasterizeSvg(
      svg, applySize ? size : existingPicture.width, getDiagramSettings(payload.settings).imageQuality,
    ))
    suppressDiagramPlaceholder(contentControl)
    const replacement = await replaceDiagramPicture(context, existingPicture, raster, payload, applySize, {
      altTextTitle: existingPicture.altTextTitle || 'Mermaid diagram',
      altTextDescription: existingPicture.altTextDescription || 'Diagram created with Mermaid Office.',
    })
    replacement.select()
    await context.sync()
  })

  return 'png'
}

async function insertDiagramPicture(
  context: Word.RequestContext,
  target: Word.Range | Word.ContentControl,
  png: string,
  options: PictureOptions,
  replacePlaceholder = false,
): Promise<Word.Range> {
  const pixels = getPngDimensions(png)
  // Keep tall/large pictures on the explicit-geometry path: Word can auto-fit
  // them after a native insertion and override the requested frame.
  const useNative = pixels.width <= 1536 && pixels.height <= 1536
    && pixels.width * pixels.height <= 2 * 1024 * 1024
    && options.width > 0 && options.width <= 468 && options.height > 0 && options.height <= 432
  if (!useNative) {
    return target.insertOoxml(createDiagramPictureOoxml(png, options), Word.InsertLocation.replace)
  }

  const insertNative = () => {
    const picture = target.insertInlinePictureFromBase64(png, Word.InsertLocation.replace)
    picture.lockAspectRatio = false
    picture.width = options.width
    picture.height = options.height
    picture.lockAspectRatio = true
    picture.altTextTitle = options.altTextTitle ?? 'Mermaid diagram'
    picture.altTextDescription = options.altTextDescription ?? 'Diagram created with Mermaid Office.'
    return picture
  }
  let picture = insertNative()
  if (replacePlaceholder) {
    // A new content control can retain its placeholder's frame on first insert.
    await context.sync()
    picture = insertNative()
  }
  return picture.getRange()
}

async function replaceDiagramPicture(
  context: Word.RequestContext,
  existingPicture: Word.InlinePicture,
  raster: RasterizedDiagram,
  payload: DiagramPayload,
  applySize: boolean,
  altText: Pick<Word.InlinePicture, 'altTextTitle' | 'altTextDescription'> = existingPicture,
): Promise<Word.Range> {
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
  // Replace only the picture: deleting its control can invalidate Word's range
  // and remove surrounding text or other diagrams.
  const replacement = await insertDiagramPicture(context, existingPicture.getRange(), png, {
    ...dimensions,
    altTextTitle: altText.altTextTitle,
    altTextDescription: altText.altTextDescription,
  })
  context.document.settings.add(getDocumentSettingKey(payload.id), JSON.stringify(payload))
  return replacement
}

export async function updateDiagramById(
  svg: string,
  existing: DiagramPayload,
  source: string,
  theme: DiagramTheme = existing.theme,
  size: DiagramSize = existing.size,
  applySize = false,
  renderedRaster?: RasterizedDiagram,
  settings?: DiagramSettings,
): Promise<DiagramFormat> {
  if (typeof Word === 'undefined') {
    throw new Error('Open Mermaid Office inside Microsoft Word to update a diagram.')
  }

  const payload: DiagramPayload = {
    ...existing, source, theme, size, format: 'png',
    ...(settings ? { settings: { ...settings } } : {}),
  }
  await Word.run(async (context) => {
    const controls = context.document.contentControls.getByTag(getContentControlTag(existing.id))
    controls.load('items')
    await context.sync()
    if (controls.items.length === 0) {
      throw new Error('The Mermaid diagram was deleted or can no longer be found.')
    }
    if (controls.items.length !== 1) {
      throw new Error('Multiple Mermaid diagrams share this ID. Select the diagram again to edit it.')
    }

    const pictures = controls.items[0].inlinePictures
    const setting = context.document.settings.getItemOrNullObject(getDocumentSettingKey(existing.id))
    pictures.load('items')
    setting.load('value')
    await context.sync()
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
        stored.format !== existing.format ||
        !sameDiagramSettings(stored.settings, existing.settings)
      ) {
        throw new Error('The Mermaid diagram changed in another editor. Select it again before updating.')
      }
    }

    const picture = pictures.items[0]
    picture.load('altTextTitle,altTextDescription,width')
    await context.sync()
    const raster = renderedRaster ?? (await rasterizeSvg(
      svg, applySize ? size : picture.width, getDiagramSettings(payload.settings).imageQuality,
    ))
    suppressDiagramPlaceholder(controls.items[0])
    await replaceDiagramPicture(context, picture, raster, payload, applySize)
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
  settings?: DiagramSettings,
): Promise<DiagramFormat> {
  return (await insertDiagramWithPayload(svg, source, theme, size, renderedRaster, {}, settings)).format
}

export async function insertDiagramWithPayload(
  svg: string,
  source: string,
  theme: DiagramTheme = 'default',
  size: DiagramSize = 'medium',
  renderedRaster?: RasterizedDiagram,
  options: DiagramInsertionOptions = {},
  settings?: DiagramSettings,
): Promise<DiagramPayload> {
  if (typeof Office === 'undefined' || !Office.context?.document) {
    throw new Error('Open Mermaid Office inside Microsoft Word to insert a diagram.')
  }

  const payload = createDiagramPayload(source, 'png', theme, size, settings)
  const raster = renderedRaster ?? (await rasterizeSvg(svg, size, getDiagramSettings(settings).imageQuality))
  const png = embedPayloadInPng(raster.base64, payload)
  await insertPngObject({ ...raster, base64: png }, payload, options)
  return payload
}
