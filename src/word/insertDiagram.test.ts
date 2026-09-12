import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  fitDiagram,
  findVisiblePixelBounds,
  insertPngObject,
  normalizeSvgDimensions,
  updateDiagram,
} from './insertDiagram'
import { createDiagramPayload, getDocumentSettingKey } from '../metadata/payload'

const transparentPixel =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+XwX9WQAAAABJRU5ErkJggg=='

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
    const insertInlinePictureFromBase64 = vi.fn()
    const contentControl = {
      tag: '',
      title: '',
      appearance: '',
      cannotDelete: true,
      cannotEdit: true,
      insertInlinePictureFromBase64,
    }
    const picture = {
      altTextTitle: '',
      altTextDescription: '',
      width: 0,
      height: 0,
    }
    insertInlinePictureFromBase64.mockReturnValue(picture)
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

    expect(insertInlinePictureFromBase64).toHaveBeenCalledTimes(2)
    expect(insertInlinePictureFromBase64).toHaveBeenNthCalledWith(
      1,
      expect.any(String),
      'Replace',
    )
    expect(insertInlinePictureFromBase64).toHaveBeenNthCalledWith(
      2,
      expect.any(String),
      'Replace',
    )
    const syncCount = enclosingKind === 'expanded' ? 9 : enclosingKind === 'empty' ? 6 : 4
    expect(sync).toHaveBeenCalledTimes(syncCount)
    expect(sync.mock.invocationCallOrder[syncCount - 3]).toBeLessThan(
      insertInlinePictureFromBase64.mock.invocationCallOrder[0],
    )
    expect(sync.mock.invocationCallOrder[syncCount - 2]).toBeLessThan(
      insertInlinePictureFromBase64.mock.invocationCallOrder[1],
    )
    expect(contentControl.tag).toBe(`mermaid-office:v1:${payload.id}`)
    expect(contentControl.appearance).toBe('Hidden')
    expect(picture.width).toBe(324)
    expect(picture.height).toBeCloseTo(79.61, 2)
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
    const insertInlinePictureFromBase64 = vi.fn().mockReturnValue({
      altTextTitle: '',
      altTextDescription: '',
      width: 0,
      height: 0,
      insertContentControl: vi.fn(),
    })
    const track = vi.fn()
    const untrack = vi.fn()
    const getRange = vi.fn().mockReturnValue({
      track,
      untrack,
      insertInlinePictureFromBase64,
    })
    const existingPicture = {
      isNullObject: false,
      width: 324,
      altTextTitle: 'Mermaid diagram',
      altTextDescription: 'Diagram created with Mermaid Office.',
      load: vi.fn(),
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
      getRange,
      inlinePictures: {
        getFirstOrNullObject: vi.fn().mockReturnValue(existingPicture),
      },
    }
    const replacementControl = {
      tag: '',
      title: '',
      appearance: '',
      cannotDelete: true,
      cannotEdit: true,
      select: vi.fn(),
      inlinePictures: {
        getFirst: vi.fn(),
      },
    }
    const replacementPicture = {
      altTextTitle: '',
      altTextDescription: '',
      width: 0,
      height: 0,
      insertContentControl: vi.fn().mockReturnValue(replacementControl),
    }
    const wrappedPicture = {
      altTextTitle: '',
      altTextDescription: '',
      width: 0,
      height: 0,
    }
    replacementControl.inlinePictures.getFirst.mockReturnValue(wrappedPicture)
    insertInlinePictureFromBase64.mockReturnValue(replacementPicture)
    const context = {
      document: {
        getSelection: () => ({
          parentContentControlOrNullObject: contentControl,
          inlinePictures: {
            getFirstOrNullObject: vi.fn().mockReturnValue({
              isNullObject: !pictureSelected,
              parentContentControlOrNullObject: contentControl,
            }),
          },
        }),
        settings: { add: vi.fn() },
      },
      sync,
    }

    vi.stubGlobal('Word', {
      run: vi.fn(async (callback) => callback(context)),
      InsertLocation: { after: 'After' },
      RangeLocation: { before: 'Before' },
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
      expect(insertInlinePictureFromBase64).not.toHaveBeenCalled()
      return
    }
    await update

    expect(getRange).toHaveBeenCalledWith('Before')
    expect(track).toHaveBeenCalledOnce()
    expect(contentControl.delete).toHaveBeenCalledWith(false)
    expect(insertInlinePictureFromBase64).toHaveBeenCalledWith(
      expect.any(String),
      'After',
    )
    expect(
      contentControl.delete.mock.invocationCallOrder[0],
    ).toBeLessThan(insertInlinePictureFromBase64.mock.invocationCallOrder[0])
    expect(replacementPicture.insertContentControl).toHaveBeenCalledOnce()
    expect(replacementControl.inlinePictures.getFirst).toHaveBeenCalledOnce()
    expect(replacementControl.tag).toBe(`mermaid-office:v1:${existing.id}`)
    expect(wrappedPicture.width).toBe(324)
    expect(wrappedPicture.height).toBeCloseTo(194.4, 2)
    expect(replacementControl.select).toHaveBeenCalledOnce()
    expect(untrack).toHaveBeenCalledOnce()
  })
})
