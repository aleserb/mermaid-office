import { embedPayloadInPng, getPngDimensions, setPngPhysicalWidth } from '../metadata/pngMetadata'
import { configureDiagramContentControl, suppressDiagramPlaceholder } from './contentControls'
import { createDiagramPictureOoxml, type PictureOptions } from './pictureOoxml'
import { getDiagramSettings, type DiagramSettings } from '../metadata/diagramSettings'
import type { InsertDiagramRequest, UpdateDiagramRequest } from '../office/diagramRequests'
import { DIAGRAM_WIDTHS, fitDiagram } from '../rendering/diagramSizing'
import { rasterizeSvg, type RasterizedDiagram } from '../rendering/rasterizeSvg'
import {
  createDiagramPayload,
  getContentControlTag,
  getDiagramIdFromTag,
  getDocumentSettingKey,
  parseDiagramPayload,
  sameDiagramPayload,
  type DiagramFormat,
  type DiagramPayload,
  type DiagramSize,
  type DiagramTheme,
} from '../metadata/payload'

export type { DiagramFormat } from '../metadata/payload'

export interface DiagramInsertionOptions {
  requireEmptySelection?: boolean
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

export async function updateDiagramById({
  svg,
  existing,
  draft: { source, theme, size, settings },
  applySize = false,
  raster: renderedRaster,
}: UpdateDiagramRequest): Promise<DiagramFormat> {
  if (typeof Word === 'undefined') {
    throw new Error('Open Mermaid Office inside Microsoft Word to update a diagram.')
  }

  const payload: DiagramPayload = {
    ...existing, source, theme, size, format: 'png',
    settings: { ...settings },
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
      if (!sameDiagramPayload(stored, existing)) {
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

export async function insertDiagram(
  svg: string,
  source: string,
  theme: DiagramTheme = 'default',
  size: DiagramSize = 'medium',
  renderedRaster?: RasterizedDiagram,
  settings?: DiagramSettings,
): Promise<DiagramFormat> {
  return (await insertDiagramWithPayload({
    svg,
    draft: { source, theme, size, settings: getDiagramSettings(settings) },
    raster: renderedRaster,
  })).format
}

export async function insertDiagramWithPayload({
  svg,
  draft: { source, theme, size, settings },
  raster: renderedRaster,
  requireEmptySelection,
}: InsertDiagramRequest): Promise<DiagramPayload> {
  if (typeof Office === 'undefined' || !Office.context?.document) {
    throw new Error('Open Mermaid Office inside Microsoft Word to insert a diagram.')
  }

  const payload = createDiagramPayload(source, 'png', theme, size, settings)
  const raster = renderedRaster ?? (await rasterizeSvg(svg, size, getDiagramSettings(settings).imageQuality))
  const png = embedPayloadInPng(raster.base64, payload)
  await insertPngObject({ ...raster, base64: png }, payload, { requireEmptySelection })
  return payload
}
