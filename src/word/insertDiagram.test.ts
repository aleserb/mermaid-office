import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  fitDiagram,
  insertPngObject,
  isSvgInsertionSupported,
} from './insertDiagram'
import { createDiagramPayload, getDocumentSettingKey } from '../metadata/payload'

describe('isSvgInsertionSupported', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
  })

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

  it('returns false outside an Office host', () => {
    vi.stubGlobal('Office', undefined)
    expect(isSvgInsertionSupported()).toBe(false)
  })

  it('checks for ImageCoercion 1.2', () => {
    const isSetSupported = vi.fn().mockReturnValue(true)
    vi.stubGlobal('Office', {
      context: { requirements: { isSetSupported } },
    })

    expect(isSvgInsertionSupported()).toBe(true)
    expect(isSetSupported).toHaveBeenCalledWith('ImageCoercion', '1.2')
  })

  it('commits a content control before inserting a PNG into it', async () => {
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

    expect(insertInlinePictureFromBase64).toHaveBeenCalledWith('png-data', 'Replace')
    expect(sync).toHaveBeenCalledTimes(2)
    expect(sync.mock.invocationCallOrder[0]).toBeLessThan(
      insertInlinePictureFromBase64.mock.invocationCallOrder[0],
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
})
