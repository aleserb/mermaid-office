import { describe, expect, it, vi } from 'vitest'
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

  it('replaces the first PNG after Word establishes its content control', async () => {
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
    const settingsAdd = vi.fn()
    const context = {
      document: {
        getSelection: () => ({ insertContentControl }),
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
      { base64: 'png-data', width: 262.5, height: 64.5 },
      payload,
    )

    expect(insertInlinePictureFromBase64).toHaveBeenCalledTimes(2)
    expect(insertInlinePictureFromBase64).toHaveBeenNthCalledWith(
      1,
      'png-data',
      'Replace',
    )
    expect(insertInlinePictureFromBase64).toHaveBeenNthCalledWith(
      2,
      'png-data',
      'Replace',
    )
    expect(sync).toHaveBeenCalledTimes(3)
    expect(sync.mock.invocationCallOrder[0]).toBeLessThan(
      insertInlinePictureFromBase64.mock.invocationCallOrder[0],
    )
    expect(sync.mock.invocationCallOrder[1]).toBeLessThan(
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
  })

  it('replaces only the existing picture range when updating a diagram', async () => {
    const sync = vi.fn().mockResolvedValue(undefined)
    const insertInlinePictureFromBase64 = vi.fn().mockReturnValue({
      altTextTitle: '',
      altTextDescription: '',
      width: 0,
      height: 0,
    })
    const getRange = vi.fn().mockReturnValue({ insertInlinePictureFromBase64 })
    const existingPicture = {
      isNullObject: false,
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
            getFirstOrNullObject: vi.fn().mockReturnValue({ isNullObject: true }),
          },
        }),
        settings: { add: vi.fn() },
      },
      sync,
    }

    vi.stubGlobal('Word', {
      run: vi.fn(async (callback) => callback(context)),
      InsertLocation: { replace: 'Replace' },
    })

    await updateDiagram(
      '<svg/>',
      existing,
      'flowchart LR\nA',
      'default',
      'medium',
      false,
      { base64: transparentPixel, width: 100, height: 60 },
    )

    expect(getRange).toHaveBeenCalledOnce()
    expect(insertInlinePictureFromBase64).toHaveBeenCalledWith(
      expect.any(String),
      'Replace',
    )
  })
})
