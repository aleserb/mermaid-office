import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  fitDiagram,
  findVisiblePixelBounds,
  insertDiagram,
  insertDiagramWithPayload,
  insertPngObject,
  normalizeSvgDimensions,
  updateDiagram,
  updateDiagramById,
} from './insertDiagram'
import { createDiagramPayload, getContentControlTag, getDocumentSettingKey } from '../metadata/payload'
import { embedPayloadInPng, readPayloadFromPng, setPngPhysicalWidth } from '../metadata/pngMetadata'

const transparentPixel =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+XwX9WQAAAABJRU5ErkJggg=='

function readPictureOoxml(xml: string) {
  const document = new DOMParser().parseFromString(xml, 'application/xml')
  expect(document.querySelector('parsererror')).toBeNull()
  const extent = document.getElementsByTagNameNS(
    'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'extent',
  )[0]
  const properties = document.getElementsByTagNameNS(
    'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'docPr',
  )[0]
  return {
    base64: document.getElementsByTagNameNS('*', 'binaryData')[0].textContent!,
    width: Number(extent.getAttribute('cx')) / 12700,
    height: Number(extent.getAttribute('cy')) / 12700,
    altTextTitle: properties.getAttribute('title'),
    altTextDescription: properties.getAttribute('descr'),
  }
}

function mockSavedDiagram() {
  const existing = createDiagramPayload('flowchart LR\nA --> B', 'png', 'forest', 'large')
  const select = vi.fn()
  const replacement = { select }
  const insertPicture = vi.fn().mockReturnValue(replacement)
  const picture = {
    width: 287,
    altTextTitle: 'Custom diagram title',
    altTextDescription: 'Accessible description',
    load: vi.fn(),
    getRange: vi.fn().mockReturnValue({ insertOoxml: insertPicture }),
  }
  const pictures = { items: [picture], load: vi.fn() }
  const control = {
    inlinePictures: pictures,
    delete: vi.fn(),
    select: vi.fn(),
    insertInlinePictureFromBase64: vi.fn(),
    insertOoxml: vi.fn(),
  }
  const controls = { items: [control], load: vi.fn() }
  const setting = { isNullObject: false, value: JSON.stringify(existing), load: vi.fn() }
  const settingsAdd = vi.fn()
  const getByTag = vi.fn().mockReturnValue(controls)
  const getSelection = vi.fn(() => {
    throw new Error('The cursor is on an unrelated diagram; selection must not be read.')
  })
  const context = {
    document: {
      getSelection,
      contentControls: { getByTag },
      settings: { add: settingsAdd, getItemOrNullObject: vi.fn().mockReturnValue(setting) },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  }
  vi.stubGlobal('Word', {
    run: vi.fn(async (callback) => callback(context)),
    InsertLocation: { replace: 'Replace' },
  })
  return {
    existing, picture, pictures, control, controls, setting, settingsAdd,
    getByTag, getSelection, insertPicture, replacement, select, context,
  }
}

describe('Word diagram insertion', () => {
  afterEach(() => vi.unstubAllGlobals())
  describe('fitDiagram', () => {
    it('keeps small diagrams at their natural size', () => {
      expect(fitDiagram(200, 100)).toEqual({ width: 200, height: 100 })
    })

    it('scales large diagrams without changing their aspect ratio', () => {
      expect(fitDiagram(1000, 500)).toEqual({ width: 500, height: 250 })
      expect(fitDiagram(500, 1300)).toEqual({ width: 250, height: 650 })
    })

    it('can upscale a diagram to a selected width preset', () => {
      expect(fitDiagram(200, 100, 400, 650, true)).toEqual({
        width: 400,
        height: 200,
      })
    })
  })

  it('finds the visible alpha bounds for PNG cropping', () => {
    const pixels = new Uint8ClampedArray(4 * 4 * 3)
    pixels[(1 * 4 + 1) * 4 + 3] = 255
    pixels[(2 * 4 + 3) * 4 + 3] = 128

    expect(findVisiblePixelBounds(pixels, 4, 3)).toEqual({
      left: 1,
      top: 1,
      width: 3,
      height: 2,
    })
    expect(findVisiblePixelBounds(new Uint8ClampedArray(16), 2, 2)).toBeNull()
  })

  it('replaces responsive Mermaid dimensions with explicit SVG bounds', () => {
    const result = normalizeSvgDimensions(
      '<svg xmlns="http://www.w3.org/2000/svg" width="100%" style="max-width: 262.5px;" viewBox="0 0 262.5 64.5"></svg>',
    )

    expect(result.width).toBe(262.5)
    expect(result.height).toBe(64.5)
    expect(result.svg).toContain('width="262.5"')
    expect(result.svg).toContain('height="64.5"')
    expect(result.svg).not.toContain('max-width')
  })

  it.each(['outside', 'expanded', 'empty', 'unrelated'])(
    'inserts independently of an %s enclosing control',
    async (enclosingKind) => {
    const sync = vi.fn().mockResolvedValue(undefined)
    const insertOoxml = vi.fn()
    const contentControl = {
      tag: '',
      title: '',
      appearance: '',
      cannotDelete: true,
      cannotEdit: true,
      insertOoxml,
    }
    const insertContentControl = vi.fn().mockReturnValue(contentControl)
    const isolated = { tag: '', title: '', appearance: '', cannotDelete: true, cannotEdit: true }
    const pictureRange = {
      track: vi.fn(),
      untrack: vi.fn(),
      insertContentControl: vi.fn().mockReturnValue(isolated),
    }
    const enclosing = {
      isNullObject: enclosingKind === 'outside',
      tag: enclosingKind === 'unrelated' ? 'other-addin' : 'mermaid-office:v1:previous',
      load: vi.fn(),
      delete: vi.fn(),
      inlinePictures: {
        getFirstOrNullObject: vi.fn().mockReturnValue({
          isNullObject: enclosingKind === 'empty',
          getRange: vi.fn().mockReturnValue(pictureRange),
        }),
      },
    }
    const settingsAdd = vi.fn()
    const context = {
      document: {
        getSelection: () => ({
          insertContentControl,
          parentContentControlOrNullObject: enclosing,
        }),
        settings: { add: settingsAdd },
      },
      sync,
    }
    const run = vi.fn(async (callback) => callback(context))

    vi.stubGlobal('Word', {
      run,
      InsertLocation: { replace: 'Replace' },
      ContentControlAppearance: { hidden: 'Hidden' },
    })

    const payload = createDiagramPayload(
      'flowchart LR\nA --> B',
      'png',
      'default',
      'medium',
    )
    await insertPngObject(
      { base64: transparentPixel, width: 262.5, height: 64.5 },
      payload,
    )

    expect(insertOoxml).toHaveBeenCalledOnce()
    expect(contentControl).toHaveProperty('placeholderText', ' ')
    if (enclosingKind === 'expanded') {
      expect(isolated).toHaveProperty('placeholderText', ' ')
    }
    expect(insertOoxml).toHaveBeenNthCalledWith(
      1,
      expect.any(String),
      'Replace',
    )
    const syncCount = enclosingKind === 'expanded' ? 8 : enclosingKind === 'empty' ? 5 : 3
    expect(sync).toHaveBeenCalledTimes(syncCount)
    expect(sync.mock.invocationCallOrder[syncCount - 2]).toBeLessThan(
      insertOoxml.mock.invocationCallOrder[0],
    )
    expect(contentControl.tag).toBe(`mermaid-office:v1:${payload.id}`)
    expect(contentControl.appearance).toBe('Hidden')
    const picture = readPictureOoxml(insertOoxml.mock.calls[0][0])
    expect(picture.width).toBe(324)
    expect(picture.height).toBeCloseTo(79.61, 2)
    expect(picture.altTextTitle).toBe('Mermaid diagram')
    expect(picture.altTextDescription).toBe('Diagram created with Mermaid Office.')
    expect(settingsAdd).toHaveBeenCalledWith(
      getDocumentSettingKey(payload.id),
      JSON.stringify(payload),
    )
    if (enclosingKind === 'expanded' || enclosingKind === 'empty') {
      expect(enclosing.delete).toHaveBeenCalledWith(true)
      expect(enclosing.delete.mock.invocationCallOrder[0]).toBeLessThan(
        insertContentControl.mock.invocationCallOrder[0],
      )
    } else {
      expect(enclosing.delete).not.toHaveBeenCalled()
    }
    if (enclosingKind === 'expanded') {
      expect(pictureRange.track).toHaveBeenCalledOnce()
      expect(pictureRange.untrack).toHaveBeenCalledOnce()
      expect(isolated.tag).toBe('mermaid-office:v1:previous')
      expect(isolated.appearance).toBe('Hidden')
      expect(pictureRange.insertContentControl.mock.invocationCallOrder[0]).toBeLessThan(
        insertContentControl.mock.invocationCallOrder[0],
      )
    } else {
      expect(pictureRange.insertContentControl).not.toHaveBeenCalled()
    }
  })

  it.each([true, false])('only updates with a selected picture: %s', async (pictureSelected) => {
    const sync = vi.fn().mockResolvedValue(undefined)
    const selectReplacement = vi.fn()
    const insertOoxml = vi.fn().mockReturnValue({ select: selectReplacement })
    const getRange = vi.fn().mockReturnValue({
      insertOoxml,
    })
    const existingPicture = {
      isNullObject: !pictureSelected,
      width: 324,
      altTextTitle: 'Mermaid diagram',
      altTextDescription: 'Diagram created with Mermaid Office.',
      load: vi.fn(),
      getRange,
    }
    const existing = createDiagramPayload(
      'flowchart LR\nA --> B',
      'png',
      'default',
      'medium',
    )
    const contentControl = {
      isNullObject: false,
      tag: `mermaid-office:v1:${existing.id}`,
      load: vi.fn(),
      delete: vi.fn(),
      select: vi.fn(),
      inlinePictures: {
        getFirstOrNullObject: vi.fn().mockReturnValue(existingPicture),
      },
    }
    const context = {
      document: {
        getSelection: () => ({
          parentContentControlOrNullObject: contentControl,
          inlinePictures: {
            getFirstOrNullObject: vi.fn().mockReturnValue(existingPicture),
          },
        }),
        settings: { add: vi.fn() },
      },
      sync,
    }

    vi.stubGlobal('Word', {
      run: vi.fn(async (callback) => callback(context)),
      InsertLocation: { replace: 'Replace' },
      ContentControlAppearance: { hidden: 'Hidden' },
    })

    const update = updateDiagram(
      '<svg/>',
      existing,
      'flowchart LR\nA',
      'default',
      'medium',
      false,
      { base64: transparentPixel, width: 100, height: 60 },
    )

    if (!pictureSelected) {
      await expect(update).rejects.toThrow('Select the Mermaid diagram')
      expect(contentControl.delete).not.toHaveBeenCalled()
      expect(insertOoxml).not.toHaveBeenCalled()
      return
    }
    await update

    expect(getRange).toHaveBeenCalledOnce()
    expect(contentControl.delete).not.toHaveBeenCalled()
    expect(contentControl.inlinePictures.getFirstOrNullObject).not.toHaveBeenCalled()
    expect(insertOoxml).toHaveBeenCalledWith(
      expect.any(String),
      'Replace',
    )
    expect(contentControl.tag).toBe(`mermaid-office:v1:${existing.id}`)
    expect(contentControl).toHaveProperty('placeholderText', ' ')
    const replacementPicture = readPictureOoxml(insertOoxml.mock.calls[0][0])
    expect(replacementPicture.width).toBe(324)
    expect(replacementPicture.height).toBeCloseTo(194.4, 2)
    expect(selectReplacement).toHaveBeenCalledOnce()
    expect(contentControl.select).not.toHaveBeenCalled()
    expect(readPayloadFromPng(replacementPicture.base64)).toEqual({
      ...existing,
      source: 'flowchart LR\nA',
    })
  })

  it.each([insertDiagramWithPayload, insertDiagram])(
    '%s preserves insertion metadata and the caller return contract',
    async (insert) => {
      const insertPicture = vi.fn()
      const control = { tag: '', insertOoxml: insertPicture }
      const settingsAdd = vi.fn()
      vi.stubGlobal('Office', { context: { document: {} } })
      vi.stubGlobal('Word', {
        run: vi.fn(async (callback) => callback({
          document: {
            getSelection: () => ({
              parentContentControlOrNullObject: { isNullObject: true, load: vi.fn() },
              insertContentControl: vi.fn().mockReturnValue(control),
            }),
            settings: { add: settingsAdd },
          },
          sync: vi.fn().mockResolvedValue(undefined),
        })),
        InsertLocation: { replace: 'Replace' },
        ContentControlAppearance: { hidden: 'Hidden' },
      })
      const result = await insert('<svg/>', 'flowchart LR\nA', undefined, undefined, {
        base64: transparentPixel, width: 100, height: 50,
      })
      const picture = readPictureOoxml(insertPicture.mock.calls[0][0])
      const payload = readPayloadFromPng(picture.base64)
      expect(payload).toMatchObject({
        source: 'flowchart LR\nA', format: 'png', theme: 'default', size: 'medium',
      })
      expect(result).toEqual(insert === insertDiagramWithPayload ? payload : 'png')
      expect(control.tag).toBe(getContentControlTag(payload!.id))
      expect(settingsAdd).toHaveBeenCalledWith(getDocumentSettingKey(payload!.id), JSON.stringify(payload))
      expect(picture.width).toBe(324)
      expect(picture.height).toBe(162)
      expect(insertPicture).toHaveBeenCalledOnce()
    },
  )

  it.each(['picture', 'text', 'cursor'])('guards pane insertion when selection is %s', async (selectionKind) => {
    const pictureSelected = selectionKind === 'picture'
    const blocked = selectionKind !== 'cursor'
    const insertPicture = vi.fn()
    const insertContentControl = vi.fn().mockReturnValue({
      insertOoxml: insertPicture,
    })
    const enclosing = { isNullObject: true, load: vi.fn(), delete: vi.fn() }
    const getSelectedPicture = vi.fn().mockReturnValue({ isNullObject: !pictureSelected })
    const getSelection = vi.fn().mockReturnValue({
      parentContentControlOrNullObject: enclosing,
      inlinePictures: { getFirstOrNullObject: getSelectedPicture },
      isEmpty: !blocked,
      load: vi.fn(),
      insertContentControl,
    })
    const settingsAdd = vi.fn()
    const sync = vi.fn().mockResolvedValue(undefined)
    vi.stubGlobal('Office', { context: { document: {} } })
    vi.stubGlobal('Word', {
      run: vi.fn(async (callback) => callback({
        document: { getSelection, settings: { add: settingsAdd } },
        sync,
      })),
      InsertLocation: { replace: 'Replace' },
      ContentControlAppearance: { hidden: 'Hidden' },
    })
    const result = insertDiagramWithPayload(
      '<svg/>', 'flowchart LR\nA', undefined, undefined,
      { base64: transparentPixel, width: 100, height: 50 },
      { requireEmptySelection: true },
    )
    if (blocked) {
      await expect(result).rejects.toThrow('without selecting text or a picture')
      expect(enclosing.load).not.toHaveBeenCalled()
      expect(enclosing.delete).not.toHaveBeenCalled()
      expect(insertContentControl).not.toHaveBeenCalled()
      expect(insertPicture).not.toHaveBeenCalled()
      expect(settingsAdd).not.toHaveBeenCalled()
    } else {
      await expect(result).resolves.toMatchObject({ source: 'flowchart LR\nA', format: 'png' })
      expect(insertPicture).toHaveBeenCalledOnce()
      expect(sync.mock.invocationCallOrder[0]).toBeLessThan(
        insertContentControl.mock.invocationCallOrder[0],
      )
    }
    expect(getSelectedPicture).toHaveBeenCalledTimes(blocked ? 1 : 2)
    expect(getSelection).toHaveBeenCalledTimes(blocked ? 1 : 2)
  })

  it.each([true, false])('rechecks the cursor after unwrapping, still empty: %s', async (empty) => {
    const pictureRange = {
      track: vi.fn(), untrack: vi.fn(),
      insertContentControl: vi.fn().mockReturnValue({}),
    }
    const enclosing = {
      isNullObject: false, tag: 'mermaid-office:v1:previous', load: vi.fn(), delete: vi.fn(),
      inlinePictures: {
        getFirstOrNullObject: () => ({ isNullObject: false, getRange: () => pictureRange }),
      },
    }
    const insertPicture = vi.fn().mockReturnValue({})
    const originalSelection = {
      isEmpty: true, load: vi.fn(),
      parentContentControlOrNullObject: enclosing,
      inlinePictures: { getFirstOrNullObject: () => ({ isNullObject: true }) },
      insertContentControl: vi.fn(() => { throw new Error('The old selection path is invalid') }),
    }
    const freshSelection = {
      isEmpty: empty, load: vi.fn(),
      inlinePictures: { getFirstOrNullObject: () => ({ isNullObject: true }) },
      insertContentControl: vi.fn().mockReturnValue({ insertOoxml: insertPicture }),
    }
    const getSelection = vi.fn().mockReturnValueOnce(originalSelection).mockReturnValue(freshSelection)
    vi.stubGlobal('Word', {
      run: vi.fn(async (callback) => callback({
        document: { getSelection, settings: { add: vi.fn() } },
        sync: vi.fn().mockResolvedValue(undefined),
      })),
      InsertLocation: { replace: 'Replace' },
      ContentControlAppearance: { hidden: 'Hidden' },
    })
    const insertion = insertPngObject(
      { base64: transparentPixel, width: 100, height: 60 },
      createDiagramPayload('flowchart LR\nNew', 'png'),
      { requireEmptySelection: true },
    )
    if (empty) {
      await insertion
    } else {
      await expect(insertion).rejects.toThrow('without selecting text or a picture')
    }
    expect(enclosing.delete).toHaveBeenCalledWith(true)
    expect(originalSelection.insertContentControl).not.toHaveBeenCalled()
    expect(freshSelection.load).toHaveBeenCalledWith('isEmpty')
    expect(freshSelection.insertContentControl).toHaveBeenCalledTimes(empty ? 1 : 0)
    expect(insertPicture).toHaveBeenCalledTimes(empty ? 1 : 0)
  })
})

describe('Word diagram updates by saved ID', () => {
  afterEach(() => vi.unstubAllGlobals())
  const raster = { base64: transparentPixel, width: 100, height: 60 }

  it.each([false, true])('targets the saved ID without touching selection (applySize=%s)', async (applySize) => {
    const mock = mockSavedDiagram()
    const source = 'flowchart LR\nB --> C'
    const result = await updateDiagramById(
      '<svg/>', mock.existing, source, undefined, undefined, applySize, raster,
    )
    const payload = { ...mock.existing, source }
    const width = applySize ? 396 : 287
    expect(result).toBe('png')
    expect(mock.getByTag).toHaveBeenCalledWith(getContentControlTag(mock.existing.id))
    expect(mock.control).toHaveProperty('placeholderText', ' ')
    expect(mock.insertPicture).toHaveBeenCalledExactlyOnceWith(
      expect.any(String), 'Replace',
    )
    const replacement = readPictureOoxml(mock.insertPicture.mock.calls[0][0])
    expect(replacement).toEqual({
      base64: setPngPhysicalWidth(embedPayloadInPng(transparentPixel, payload), width),
      width, height: width * 0.6,
      altTextTitle: mock.picture.altTextTitle,
      altTextDescription: mock.picture.altTextDescription,
    })
    expect(mock.settingsAdd).toHaveBeenCalledWith(getDocumentSettingKey(mock.existing.id), JSON.stringify(payload))
    expect(mock.getSelection).not.toHaveBeenCalled()
    expect(mock.select).not.toHaveBeenCalled()
    expect(mock.control.select).not.toHaveBeenCalled()
    expect(mock.control.delete).not.toHaveBeenCalled()
    expect(mock.control.insertInlinePictureFromBase64).not.toHaveBeenCalled()
    expect(mock.control.insertOoxml).not.toHaveBeenCalled()
  })

  it.each([
    ['missing control', 'deleted'],
    ['duplicate control', 'Multiple Mermaid diagrams'],
    ['missing picture', 'exactly one picture'],
    ['ambiguous picture', 'exactly one picture'],
    ['changed source', 'changed in another editor'],
    ['changed theme', 'changed in another editor'],
    ['changed size', 'changed in another editor'],
    ['invalid metadata', 'not valid JSON'],
  ])('rejects %s without writes', async (scenario, message) => {
    const mock = mockSavedDiagram()
    if (scenario === 'missing control') mock.controls.items = []
    if (scenario === 'duplicate control') mock.controls.items.push(mock.control)
    if (scenario === 'missing picture') mock.pictures.items = []
    if (scenario === 'ambiguous picture') mock.pictures.items.push(mock.picture)
    if (scenario === 'changed source') mock.setting.value = JSON.stringify({ ...mock.existing, source: 'flowchart LR\nZ' })
    if (scenario === 'changed theme') mock.setting.value = JSON.stringify({ ...mock.existing, theme: 'dark' })
    if (scenario === 'changed size') mock.setting.value = JSON.stringify({ ...mock.existing, size: 'small' })
    if (scenario === 'invalid metadata') mock.setting.value = '{'
    await expect(updateDiagramById('<svg/>', mock.existing, 'flowchart LR\nC', undefined, undefined, false, raster))
      .rejects.toThrow(message)
    expect(mock.insertPicture).not.toHaveBeenCalled()
    expect(mock.settingsAdd).not.toHaveBeenCalled()
    expect(mock.control.delete).not.toHaveBeenCalled()
    expect(mock.control.insertInlinePictureFromBase64).not.toHaveBeenCalled()
    expect(mock.control.insertOoxml).not.toHaveBeenCalled()
    expect(mock.getSelection).not.toHaveBeenCalled()
    expect(mock.select).not.toHaveBeenCalled()
  })

  it('updates theme and size metadata and restores missing document settings', async () => {
    const mock = mockSavedDiagram()
    mock.setting.isNullObject = true
    await updateDiagramById('<svg/>', mock.existing, 'flowchart LR\nC', 'dark', 'small', true, raster)
    const replacement = readPictureOoxml(mock.insertPicture.mock.calls[0][0])
    const payload = readPayloadFromPng(replacement.base64)
    expect(payload).toEqual({ ...mock.existing, source: 'flowchart LR\nC', theme: 'dark', size: 'small' })
    expect(replacement.width).toBe(216)
    expect(mock.settingsAdd).toHaveBeenCalledWith(getDocumentSettingKey(mock.existing.id), JSON.stringify(payload))
  })

  it('preserves deliberately empty alt text and manual width when changing the size preset', async () => {
    const mock = mockSavedDiagram()
    mock.picture.altTextTitle = ''
    mock.picture.altTextDescription = ''
    await updateDiagramById('<svg/>', mock.existing, 'flowchart LR\nC', 'dark', 'small', false, raster)
    const replacement = readPictureOoxml(mock.insertPicture.mock.calls[0][0])
    expect(replacement).toMatchObject({
      width: 287, height: 287 * 0.6, altTextTitle: '', altTextDescription: '',
    })
    expect(readPayloadFromPng(replacement.base64)).toMatchObject({
      theme: 'dark', size: 'small', source: 'flowchart LR\nC', format: 'png',
    })

  })

  it('surfaces a failed OOXML replacement without deleting the diagram control', async () => {
    const mock = mockSavedDiagram()
    mock.context.sync
      .mockResolvedValueOnce(undefined)
      .mockResolvedValueOnce(undefined)
      .mockResolvedValueOnce(undefined)
      .mockRejectedValueOnce(new Error('Word could not replace the picture'))
    await expect(updateDiagramById('<svg/>', mock.existing, 'flowchart LR\nC', undefined, undefined, false, raster))
      .rejects.toThrow('Word could not replace the picture')
    expect(mock.control.delete).not.toHaveBeenCalled()
    expect(mock.select).not.toHaveBeenCalled()
  })
})
