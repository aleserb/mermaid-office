import { expect, it } from 'vitest'
import manifestSource from '../manifest.xml?raw'

it('exposes one Mermaid ribbon button for Word and Excel that opens the default pane', () => {
  const manifest = new DOMParser().parseFromString(
    manifestSource,
    'application/xml',
  )
  const controls = manifest.getElementsByTagNameNS('*', 'Control')
  expect(controls).toHaveLength(2)
  for (const control of Array.from(controls)) {
    const labelId = control.getElementsByTagNameNS('*', 'Label')[0].getAttribute('resid')
    const label = Array.from(manifest.getElementsByTagNameNS('*', 'String'))
      .find((item) => item.getAttribute('id') === labelId)
    expect(label?.getAttribute('DefaultValue')).toBe('Mermaid')
  }

  const actions = manifest.getElementsByTagNameNS('*', 'Action')
  expect(actions).toHaveLength(2)
  const defaultUrl = manifest.getElementsByTagNameNS('*', 'DefaultSettings')[0]
    .getElementsByTagNameNS('*', 'SourceLocation')[0].getAttribute('DefaultValue')
  for (const action of Array.from(actions)) {
    expect(action.getAttributeNS('http://www.w3.org/2001/XMLSchema-instance', 'type'))
      .toBe('ShowTaskpane')
    const urlId = action.getElementsByTagNameNS('*', 'SourceLocation')[0].getAttribute('resid')
    const url = Array.from(manifest.getElementsByTagNameNS('*', 'Url'))
      .find((item) => item.getAttribute('id') === urlId)?.getAttribute('DefaultValue')
    expect(url).toBe(defaultUrl)
    expect(url).toMatch(/^https:\/\/aleserb\.github\.io\/mermaid-office\/\?v=/)
  }
  expect(Array.from(manifest.getElementsByTagNameNS('*', 'Host'))
    .map((host) => host.getAttribute('Name') ?? host.getAttributeNS(
      'http://www.w3.org/2001/XMLSchema-instance',
      'type',
    ))).toEqual(['Document', 'Workbook', 'Document', 'Workbook'])
  expect(manifest.getElementsByTagNameNS('*', 'FunctionFile')).toHaveLength(0)
  expect(manifest.getElementsByTagNameNS('*', 'FunctionName')).toHaveLength(0)
})

it.each([
  ['IconUrl', 32],
  ['HighResolutionIconUrl', 64],
] as const)('references the required %s asset size', (elementName, size) => {
  const manifest = new DOMParser().parseFromString(manifestSource, 'application/xml')
  const url = new URL(manifest.getElementsByTagNameNS('*', elementName)[0].getAttribute('DefaultValue')!)
  expect(url.origin).toBe('https://aleserb.github.io')
  expect(url.pathname).toBe(`/mermaid-office/assets/icon-${size}.png`)
})
