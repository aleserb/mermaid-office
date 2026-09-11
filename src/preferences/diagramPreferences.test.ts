import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  getPreferredTheme,
  setPreferredTheme,
} from './diagramPreferences'

describe('diagram preferences', () => {
  afterEach(() => {
    window.localStorage.clear()
    vi.restoreAllMocks()
  })

  it('persists valid theme values', () => {
    setPreferredTheme('redux-color')

    expect(getPreferredTheme()).toBe('redux-color')
  })

  it('uses safe defaults when storage is unavailable', () => {
    vi.spyOn(Storage.prototype, 'getItem').mockImplementation(() => {
      throw new DOMException('Blocked', 'SecurityError')
    })
    vi.spyOn(console, 'warn').mockImplementation(() => undefined)

    expect(getPreferredTheme()).toBe('default')
  })
})
