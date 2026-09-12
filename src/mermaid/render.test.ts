import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { MermaidConfig } from 'mermaid'
import { DEFAULT_DIAGRAM_SETTINGS, type DiagramSettings } from '../metadata/diagramSettings'
import type { DiagramTheme } from '../metadata/payload'

const mocks = vi.hoisted(() => ({
  initialize: vi.fn<(config: MermaidConfig) => void>(),
  render: vi.fn<(id: string, source: string) => Promise<{ svg: string }>>(),
  registerLayoutLoaders: vi.fn(),
  loadElk: vi.fn(),
  layouts: [{ name: 'elk', loader: vi.fn() }],
}))

vi.mock('mermaid', () => ({
  default: {
    initialize: mocks.initialize,
    render: mocks.render,
    registerLayoutLoaders: mocks.registerLayoutLoaders,
  },
}))
function settings(overrides: Partial<DiagramSettings>): DiagramSettings {
  return { ...DEFAULT_DIAGRAM_SETTINGS, ...overrides }
}

function lastConfig(): MermaidConfig {
  return mocks.initialize.mock.calls.at(-1)![0]
}

beforeEach(() => {
  vi.resetModules()
  vi.resetAllMocks()
  vi.doMock('@mermaid-js/layout-elk', () => {
    mocks.loadElk()
    return { default: mocks.layouts }
  })
  mocks.render.mockResolvedValue({ svg: '<svg xmlns="http://www.w3.org/2000/svg"><text>Diagram</text></svg>' })
})

describe('detectDiagramKind', () => {
  it.each([
    ['flowchart LR\nA --> B', 'flowchart'],
    [' \r\n\tgraph TD; A-->B', 'flowchart'],
    ['flowchart-elk LR\nA-->B', 'flowchart'],
    ['sequenceDiagram\nA->>B: Hello', 'sequence'],
    ['classDiagram\nclass A', 'class'],
    ['classDiagram-v2\nclass A', 'class'],
    ['erDiagram\nA ||--o{ B : has', 'er'],
    ['stateDiagram\nA --> B', 'state'],
    ['stateDiagram-v2\nA --> B', 'state'],
    ['%% flowchart LR\n%% second comment\nsequenceDiagram', 'sequence'],
    ['---\ntitle: flowchart LR\nconfig:\n  theme: dark\n---\nclassDiagram', 'class'],
    [' \uFEFF---\r\ntitle: Test\r\n---\r\n%% comment\r\nstateDiagram-v2', 'state'],
    ['%%{init: {"theme":"dark"}}%%\nflowchart LR', 'flowchart'],
    ['%%{init: {\n"sequence": {"wrap": true}\n}}%%\nsequenceDiagram', 'sequence'],
    ['%% comment\n---\ntitle: Test\n---\n%%{init: {}}%%\nerDiagram', 'er'],
    ['%%{initialize: {}}%% %%{init: {}}%%\nflowchart LR', 'flowchart'],
    ['pie\n"flowchart LR": 10', 'other'],
    ['A["sequenceDiagram"]', 'other'],
    ['flowchartOther LR', 'other'],
    ['sequenceDiagram-invalid', 'other'],
    ['%% sequenceDiagram', 'other'],
    ['%%{init: {}\nflowchart LR', 'other'],
    ['---\ntitle: Test\nflowchart LR', 'other'],
    ['', 'other'],
  ])('detects only the header in %j', async (source, expected) => {
    const { detectDiagramKind } = await import('./render')
    expect(detectDiagramKind(source)).toBe(expected)
    expect(mocks.loadElk).not.toHaveBeenCalled()
  })
})

describe('renderMermaid', () => {
  it('preserves the default flowchart configuration and sanitizes SVG', async () => {
    mocks.render.mockResolvedValue({
      svg: '<svg xmlns="http://www.w3.org/2000/svg" onload="alert(1)"><style>text{fill:red}</style><script>alert(1)</script><text>Safe</text><foreignObject>HTML</foreignObject></svg>',
    })
    const { renderMermaid } = await import('./render')
    const svg = await renderMermaid('flowchart LR\nA-->B')
    expect(lastConfig()).toEqual({
      startOnLoad: false,
      securityLevel: 'strict',
      theme: 'default',
      look: 'classic',
      htmlLabels: false,
      suppressErrorRendering: true,
      secure: [
        'secure', 'securityLevel', 'startOnLoad', 'maxTextSize',
        'suppressErrorRendering', 'maxEdges', 'htmlLabels', 'theme', 'themeVariables', 'look',
      ],
    })
    expect(svg).toContain('<text>Safe</text>')
    expect(svg).toContain('<style>')
    expect(svg).not.toMatch(/onload|script|foreignObject/)
    expect(mocks.loadElk).not.toHaveBeenCalled()
  })

  it.each<DiagramTheme>(['default', 'dark', 'forest', 'neutral', 'neo', 'neo-dark', 'redux', 'redux-dark-color'])(
    'uses the existing automatic look for %s', async theme => {
      const { renderMermaid } = await import('./render')
      await renderMermaid('flowchart LR\nA-->B', theme)
      expect(lastConfig().look).toBe(/^(neo|redux)/.test(theme) ? 'neo' : 'classic')
    },
  )

  it.each<DiagramSettings['look']>(['classic', 'neo', 'handDrawn'])('maps explicit look %s', async look => {
    const { renderMermaid } = await import('./render')
    await renderMermaid('classDiagram\nclass A', 'neo', settings({ look }))
    expect(lastConfig().look).toBe(look)
    expect(lastConfig().handDrawnSeed).toBe(look === 'handDrawn' ? 1 : undefined)
  })

  it('maps fonts to theme variables and all sequence label font options', async () => {
    const { renderMermaid } = await import('./render')
    await renderMermaid('sequenceDiagram\nA->>B: Hello', 'dark', settings({
      fontFamily: 'Times New Roman',
      fontSize: 20,
      sequenceNumbers: true,
      sequenceWrap: true,
      sequenceMirrorActors: false,
    }))
    expect(lastConfig()).toMatchObject({
      fontFamily: 'Times New Roman',
      fontSize: 20,
      themeVariables: { fontFamily: 'Times New Roman', fontSize: '20px' },
      sequence: {
        actorFontFamily: 'Times New Roman',
        noteFontFamily: 'Times New Roman',
        messageFontFamily: 'Times New Roman',
        actorFontSize: 20,
        noteFontSize: 20,
        messageFontSize: 20,
        showSequenceNumbers: true,
        wrap: true,
        mirrorActors: false,
      },
    })
    expect(lastConfig().secure).toEqual(expect.arrayContaining([
      'fontFamily', 'fontSize', 'actorFontFamily', 'noteFontFamily', 'messageFontFamily',
      'actorFontSize', 'noteFontSize', 'messageFontSize', 'showSequenceNumbers', 'wrap', 'mirrorActors',
    ]))
    expect(lastConfig().secure).not.toContain('sequence')
  })

  it('preserves default sequence options without fixing theme fonts or spacing', async () => {
    const { renderMermaid } = await import('./render')
    const source = 'sequenceDiagram\nautonumber\nA->>B: Hello'
    await renderMermaid(source)
    expect(lastConfig()).not.toHaveProperty('sequence')
    for (const key of ['showSequenceNumbers', 'wrap', 'mirrorActors']) {
      expect(lastConfig().secure).not.toContain(key)
    }
    expect(lastConfig()).not.toHaveProperty('fontFamily')
    expect(lastConfig()).not.toHaveProperty('fontSize')
    expect(lastConfig()).not.toHaveProperty('themeVariables')
    expect(mocks.render).toHaveBeenCalledWith(expect.any(String), source)
  })

  it.each([
    ['compact', 25, 25, 20],
    ['spacious', 80, 80, 55],
  ] as const)('maps %s spacing for flowcharts and sequences', async (spacing, nodeSpacing, actorMargin, messageMargin) => {
    const { renderMermaid } = await import('./render')
    await renderMermaid('flowchart LR\nA-->B', 'default', settings({ spacing }))
    expect(lastConfig().flowchart).toEqual({ nodeSpacing, rankSpacing: nodeSpacing })
    expect(lastConfig().secure).toEqual(expect.arrayContaining(['nodeSpacing', 'rankSpacing']))
    expect(lastConfig().secure).not.toContain('flowchart')
    await renderMermaid('sequenceDiagram\nA->>B: Hello', 'default', settings({ spacing }))
    expect(lastConfig().sequence).toMatchObject({ actorMargin, messageMargin })
    expect(lastConfig().secure).toEqual(expect.arrayContaining(['actorMargin', 'messageMargin']))
  })

  it.each<DiagramSettings['curve']>(['linear', 'basis', 'step'])('maps flowchart curve %s', async curve => {
    const { renderMermaid } = await import('./render')
    await renderMermaid('graph TD\nA-->B', 'default', settings({ curve }))
    expect(lastConfig().flowchart?.curve).toBe(curve)
    expect(lastConfig().secure).toContain('curve')
  })

  it('uses the modern Dagre renderer without loading ELK', async () => {
    const { renderMermaid } = await import('./render')
    await renderMermaid('graph TD\nA-->B', 'default', settings({ layout: 'dagre' }))
    expect(lastConfig()).toMatchObject({ layout: 'dagre', flowchart: { defaultRenderer: 'dagre-wrapper' } })
    expect(lastConfig().secure).toEqual(expect.arrayContaining(['layout', 'defaultRenderer']))
    expect(mocks.loadElk).not.toHaveBeenCalled()
  })

  it('lazily registers ELK once for requested flowchart layouts', async () => {
    const { renderMermaid } = await import('./render')
    await renderMermaid('flowchart LR\nA-->B')
    expect(mocks.loadElk).not.toHaveBeenCalled()
    await renderMermaid('flowchart LR\nA-->B', 'default', settings({ layout: 'elk' }))
    await renderMermaid('graph TD\nA-->B', 'dark', settings({ layout: 'elk' }))
    expect(mocks.loadElk).toHaveBeenCalledTimes(1)
    expect(mocks.registerLayoutLoaders).toHaveBeenCalledExactlyOnceWith(mocks.layouts)
    expect(lastConfig()).toMatchObject({ layout: 'elk', flowchart: { defaultRenderer: 'dagre-wrapper' } })
  })

  it('surfaces ELK import failures rather than falling back', async () => {
    mocks.loadElk.mockImplementation(() => { throw new Error('ELK download failed') })
    const { renderMermaid } = await import('./render')
    await expect(renderMermaid('flowchart LR\nA-->B', 'default', settings({ layout: 'elk' })))
      .rejects.toThrow()
    expect(mocks.initialize).not.toHaveBeenCalled()
    expect(mocks.render).not.toHaveBeenCalled()
    await renderMermaid('flowchart LR\nA-->B')
    expect(mocks.render).toHaveBeenCalledTimes(1)
  })

  it('surfaces ELK renderer failures without retrying with Dagre', async () => {
    mocks.render.mockRejectedValueOnce(new Error('ELK engine failed'))
    const { renderMermaid } = await import('./render')
    await expect(renderMermaid('flowchart LR\nA-->B', 'default', settings({ layout: 'elk' })))
      .rejects.toThrow('ELK engine failed')
    expect(mocks.render).toHaveBeenCalledTimes(1)
  })

  it.each(['classDiagram', 'erDiagram', 'stateDiagram-v2', 'pie', 'sequenceDiagram'])(
    'does not configure flowchart-only options or load ELK for %s', async source => {
      const { renderMermaid } = await import('./render')
      await renderMermaid(source, 'default', settings({ layout: 'elk', spacing: 'compact', curve: 'step' }))
      expect(lastConfig()).not.toHaveProperty('layout')
      expect(lastConfig()).not.toHaveProperty('flowchart')
      if (source !== 'sequenceDiagram') expect(lastConfig()).not.toHaveProperty('sequence')
      expect(mocks.loadElk).not.toHaveBeenCalled()
    },
  )

  it('ignores image quality in Mermaid configuration', async () => {
    const { renderMermaid } = await import('./render')
    await renderMermaid('flowchart LR\nA-->B')
    const original = lastConfig()
    await renderMermaid('flowchart LR\nA-->B', 'default', settings({ imageQuality: 'high' }))
    expect(lastConfig()).toEqual(original)
  })

  it('serializes initialization with rendering and continues after failed renders', async () => {
    let rejectFirst!: (error: Error) => void
    mocks.render.mockImplementationOnce(() => new Promise((_resolve, reject) => { rejectFirst = reject }))
    const { renderMermaid } = await import('./render')
    const first = renderMermaid('flowchart LR\nA-->B', 'dark', settings({ fontSize: 12 }))
    const firstRejection = expect(first).rejects.toThrow('Invalid diagram')
    await vi.waitFor(() => expect(mocks.render).toHaveBeenCalledTimes(1))
    const second = renderMermaid('flowchart LR\nB-->C', 'forest', settings({ fontSize: 24 }))
    await Promise.resolve()
    expect(mocks.initialize).toHaveBeenCalledTimes(1)
    expect(lastConfig()).toMatchObject({ theme: 'dark', fontSize: 12 })
    rejectFirst(new Error('Invalid diagram'))
    await firstRejection
    await second
    expect(lastConfig()).toMatchObject({ theme: 'forest', fontSize: 24 })
    expect(mocks.render).toHaveBeenCalledTimes(2)
    expect(mocks.render.mock.calls[0][0]).not.toBe(mocks.render.mock.calls[1][0])
  })

  it('rejects blank source without initializing or loading layouts', async () => {
    const { renderMermaid } = await import('./render')
    await expect(renderMermaid(' \n\t')).rejects.toThrow('Enter Mermaid diagram source.')
    expect(mocks.initialize).not.toHaveBeenCalled()
    expect(mocks.loadElk).not.toHaveBeenCalled()
  })
})
