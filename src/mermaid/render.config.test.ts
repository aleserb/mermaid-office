import { afterEach, describe, expect, it, vi } from 'vitest'
import mermaid, { type MermaidConfig } from 'mermaid'
import { DEFAULT_DIAGRAM_SETTINGS } from '../metadata/diagramSettings'
import { renderMermaid } from './render'

afterEach(() => vi.restoreAllMocks())

function captureParsedConfig() {
  let config: MermaidConfig = {}
  vi.spyOn(mermaid, 'render').mockImplementation(async (_id, source) => {
    await mermaid.parse(source)
    config = mermaid.mermaidAPI.getConfig()
    return { svg: '<svg xmlns="http://www.w3.org/2000/svg"></svg>', diagramType: 'test' }
  })
  return () => config
}

describe('settings precedence with installed Mermaid', () => {
  it('protects selected flowchart settings and security without discarding unrelated source config', async () => {
    const config = captureParsedConfig()
    await renderMermaid(`---
config:
  theme: dark
  look: handDrawn
  securityLevel: loose
  htmlLabels: true
  layout: elk
  fontFamily: monospace
  fontSize: 12
  flowchart:
    curve: linear
    nodeSpacing: 2
    rankSpacing: 3
    diagramPadding: 30
    defaultRenderer: elk
---
flowchart LR
A --> B`, 'forest', {
      ...DEFAULT_DIAGRAM_SETTINGS,
      fontFamily: 'Arial',
      fontSize: 20,
      spacing: 'compact',
      curve: 'step',
      layout: 'dagre',
    })
    expect(config()).toMatchObject({
      theme: 'forest',
      look: 'classic',
      securityLevel: 'strict',
      htmlLabels: false,
      layout: 'dagre',
      fontFamily: 'Arial',
      fontSize: 20,
      themeVariables: { fontFamily: 'Arial', fontSize: '20px' },
      flowchart: {
        curve: 'step',
        nodeSpacing: 25,
        rankSpacing: 25,
        diagramPadding: 30,
        defaultRenderer: 'dagre-wrapper',
      },
    })
  })

  it('protects sequence options and fonts from init directives while allowing other sequence config', async () => {
    const config = captureParsedConfig()
    await renderMermaid(`%%{init: {"sequence": {
      "showSequenceNumbers": false, "wrap": false, "mirrorActors": true,
      "actorFontSize": 12, "noteFontSize": 12, "messageFontSize": 12,
      "actorFontFamily": "Arial", "noteFontFamily": "Arial", "messageFontFamily": "Arial",
      "actorMargin": 1, "messageMargin": 1, "diagramMarginX": 40
    }}}%%
sequenceDiagram
A->>B: Hello`, 'default', {
      ...DEFAULT_DIAGRAM_SETTINGS,
      fontFamily: 'monospace',
      fontSize: 24,
      sequenceNumbers: true,
      sequenceWrap: true,
      sequenceMirrorActors: false,
      spacing: 'spacious',
    })
    expect(config().sequence).toMatchObject({
      showSequenceNumbers: true,
      wrap: true,
      mirrorActors: false,
      actorFontFamily: 'monospace',
      noteFontFamily: 'monospace',
      messageFontFamily: 'monospace',
      actorFontSize: 24,
      noteFontSize: 24,
      messageFontSize: 24,
      actorMargin: 80,
      messageMargin: 55,
      diagramMarginX: 40,
    })
  })

  it('leaves default spacing and curve open to source configuration', async () => {
    const config = captureParsedConfig()
    await renderMermaid(`%%{init: {"flowchart": {"curve": "linear", "nodeSpacing": 35, "rankSpacing": 65}}}%%
flowchart LR
A --> B`)
    expect(config().flowchart).toMatchObject({ curve: 'linear', nodeSpacing: 35, rankSpacing: 65 })
  })

  it.each([
    `---
config:
  sequence:
    showSequenceNumbers: true
    wrap: true
    mirrorActors: false
---
sequenceDiagram
A->>B: Hello`,
    `%%{init: {"sequence": {"showSequenceNumbers": true, "wrap": true, "mirrorActors": false}}}%%
sequenceDiagram
A->>B: Hello`,
  ])('preserves legacy source sequence options with default settings', async source => {
    const config = captureParsedConfig()
    await renderMermaid(source, 'default', DEFAULT_DIAGRAM_SETTINGS)
    expect(config().sequence).toMatchObject({
      showSequenceNumbers: true,
      wrap: true,
      mirrorActors: false,
    })
  })

  it('protects only non-default sequence options', async () => {
    const config = captureParsedConfig()
    await renderMermaid(`%%{init: {"sequence": {"showSequenceNumbers": false, "wrap": true, "mirrorActors": false}}}%%
sequenceDiagram
A->>B: Hello`, 'default', { ...DEFAULT_DIAGRAM_SETTINGS, sequenceNumbers: true })
    expect(config().sequence).toMatchObject({
      showSequenceNumbers: true,
      wrap: true,
      mirrorActors: false,
    })
  })
})
