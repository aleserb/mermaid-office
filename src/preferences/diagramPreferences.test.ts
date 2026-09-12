import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  getPreferredTheme,
  setPreferredTheme,
  getPreferredSettings,
  setPreferredSettings,
} from './diagramPreferences'
import { DIAGRAM_THEMES } from '../metadata/payload'
import { DEFAULT_DIAGRAM_SETTINGS } from '../metadata/diagramSettings'

describe('diagram preferences', () => {
  afterEach(() => {
    window.localStorage.clear()
    vi.restoreAllMocks()
  })

  it('defaults to Redux Color for a new user', () => {
    expect(getPreferredTheme()).toBe('redux-color')
  })

  it('remembers applied settings for new diagrams', () => {
    expect(getPreferredSettings()).toEqual(DEFAULT_DIAGRAM_SETTINGS)
    const settings = { ...DEFAULT_DIAGRAM_SETTINGS, imageQuality: 'high' as const, fontSize: 18 as const }
    setPreferredSettings(settings)
    expect(getPreferredSettings()).toEqual(settings)
  })

  it.each(['{', '{"imageQuality":"unlimited"}', 'null'])('reports an invalid saved settings preference: %s', value => {
    localStorage.setItem('mermaid-office:preferred-settings', value)
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    expect(getPreferredSettings()).toEqual(DEFAULT_DIAGRAM_SETTINGS)
    expect(warn).toHaveBeenCalledOnce()
  })

  it.each(DIAGRAM_THEMES)('preserves the saved %s theme', (theme) => {
    setPreferredTheme(theme)

    expect(getPreferredTheme()).toBe(theme)
  })

  it('uses Redux Color when the saved preference is invalid', () => {
    window.localStorage.setItem('mermaid-office:preferred-theme', 'unknown')

    expect(getPreferredTheme()).toBe('redux-color')
  })

  it('uses safe defaults when storage is unavailable', () => {
    vi.spyOn(Storage.prototype, 'getItem').mockImplementation(() => {
      throw new DOMException('Blocked', 'SecurityError')
    })
    vi.spyOn(console, 'warn').mockImplementation(() => undefined)

    expect(getPreferredTheme()).toBe('redux-color')
    expect(getPreferredSettings()).toEqual(DEFAULT_DIAGRAM_SETTINGS)
  })
})
