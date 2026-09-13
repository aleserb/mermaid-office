import { expect, it } from 'vitest'
import { getHostAdapter } from './hostAdapter'

it('routes supported Office hosts to their document adapter', () => {
  expect(getHostAdapter().appName).toBe('Word')
  expect(getHostAdapter('Word').appName).toBe('Word')
  expect(getHostAdapter('Excel').appName).toBe('Excel')
})

it('rejects unsupported Office hosts', () => {
  expect(() => getHostAdapter('PowerPoint')).toThrow(
    'Mermaid Office supports Microsoft Word and Excel.',
  )
})
