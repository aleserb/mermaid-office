import { expect, it } from 'vitest'
import manifestSource from '../manifest.xml?raw'

it('exposes exactly one Mermaid ribbon button that opens the default pane', () => {
  const manifest = new DOMParser().parseFromString(
    manifestSource,
    'application/xml',
  )
  const controls = manifest.getElementsByTagNameNS('*', 'Control')
  expect(controls).toHaveLength(1)
  const labelId = controls[0].getElementsByTagNameNS('*', 'Label')[0].getAttribute('resid')
  const label = Array.from(manifest.getElementsByTagNameNS('*', 'String'))
    .find((item) => item.getAttribute('id') === labelId)
  expect(label?.getAttribute('DefaultValue')).toBe('Mermaid')

  const actions = manifest.getElementsByTagNameNS('*', 'Action')
  expect(actions).toHaveLength(1)
  expect(actions[0].getAttributeNS('http://www.w3.org/2001/XMLSchema-instance', 'type'))
    .toBe('ShowTaskpane')
  const urlId = actions[0].getElementsByTagNameNS('*', 'SourceLocation')[0].getAttribute('resid')
  const url = Array.from(manifest.getElementsByTagNameNS('*', 'Url'))
    .find((item) => item.getAttribute('id') === urlId)?.getAttribute('DefaultValue')
  const defaultUrl = manifest.getElementsByTagNameNS('*', 'DefaultSettings')[0]
    .getElementsByTagNameNS('*', 'SourceLocation')[0].getAttribute('DefaultValue')
  expect(url).toBe(defaultUrl)
  expect(url).toMatch(/^https:\/\/aleserb\.github\.io\/mermaid-office\/\?v=/)
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
