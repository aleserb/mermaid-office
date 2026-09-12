import { afterEach, describe, expect, it, vi } from 'vitest'
import { getSelectedDiagram } from './selection'
import { createDiagramPayload, getContentControlTag, getDocumentSettingKey } from '../metadata/payload'
import { readPayloadFromImage } from '../metadata/imageMetadata'

vi.mock('../metadata/imageMetadata', () => ({
  readPayloadFromImage: vi.fn(),
}))

afterEach(() => {
  vi.unstubAllGlobals()
  vi.resetAllMocks()
})

function mockSelection(picture: object, parent: object = { isNullObject: true }) {
  const settings = { getItemOrNullObject: vi.fn(), add: vi.fn() }
  const contentControls = { getByTag: vi.fn() }
  vi.stubGlobal('Word', {
    run: async (callback: (context: object) => Promise<unknown>) => callback({
      sync: vi.fn().mockResolvedValue(undefined),
      document: {
        getSelection: () => ({
          parentContentControlOrNullObject: parent,
          inlinePictures: { getFirstOrNullObject: () => picture },
        }),
        settings,
        contentControls,
      },
    }),
    ContentControlAppearance: { hidden: 'Hidden' },
  })
  return { settings, contentControls }
}

describe('selected diagram detection', () => {
  it.each([true, false])('returns Insert mode without a picture, parent is null: %s', async (isNullObject) => {
    const { settings } = mockSelection(
      { isNullObject: true },
      { isNullObject, tag: 'mermaid-office:v1:previous' },
    )
    expect(await getSelectedDiagram()).toBeNull()
    expect(settings.getItemOrNullObject).not.toHaveBeenCalled()
    expect(readPayloadFromImage).not.toHaveBeenCalled()
  })

  it('uses the selected picture control, not an enclosing older diagram', async () => {
    const payload = createDiagramPayload('flowchart LR\nA-->B', 'png', 'forest')
    const control = {
      isNullObject: false,
      tag: getContentControlTag(payload.id),
      load: vi.fn(),
    }
    const { settings, contentControls } = mockSelection(
      { isNullObject: false, parentContentControlOrNullObject: control },
      { isNullObject: false, tag: 'mermaid-office:v1:previous' },
    )
    settings.getItemOrNullObject.mockReturnValue({
      isNullObject: false,
      value: JSON.stringify(payload),
      load: vi.fn(),
    })
    contentControls.getByTag.mockReturnValue({ items: [control], load: vi.fn() })
    expect(await getSelectedDiagram()).toEqual(payload)
    expect(settings.getItemOrNullObject).toHaveBeenCalledWith(getDocumentSettingKey(payload.id))
  })

  it('does not treat an ordinary picture as a Mermaid diagram', async () => {
    mockSelection({
      isNullObject: false,
      parentContentControlOrNullObject: { isNullObject: true, load: vi.fn() },
      getBase64ImageSrc: () => ({ value: 'ordinary-image' }),
    })
    vi.mocked(readPayloadFromImage).mockResolvedValue(null)
    expect(await getSelectedDiagram()).toBeNull()
  })

  it('still recovers Mermaid metadata from a copied picture without a control', async () => {
    const payload = createDiagramPayload('flowchart LR\nA-->B', 'png', 'redux-color')
    const recoveredControl = { tag: '', select: vi.fn() }
    const { settings } = mockSelection({
      isNullObject: false,
      parentContentControlOrNullObject: { isNullObject: true, load: vi.fn() },
      getBase64ImageSrc: () => ({ value: 'mermaid-image' }),
      insertContentControl: () => recoveredControl,
    })
    vi.mocked(readPayloadFromImage).mockResolvedValue(payload)
    const recovered = await getSelectedDiagram()
    expect(recovered).toMatchObject({ source: payload.source, theme: payload.theme })
    expect(recovered?.id).not.toBe(payload.id)
    expect(recoveredControl.tag).toBe(getContentControlTag(recovered!.id))
    expect(settings.add).toHaveBeenCalledOnce()
  })
})
