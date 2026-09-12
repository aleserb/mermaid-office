import { afterEach, describe, expect, it, vi } from 'vitest'
import { officeUsesDarkTheme, resolveOfficeDarkTheme } from './officeTheme'

afterEach(() => vi.unstubAllGlobals())

describe('Office UI theme detection', () => {
  it('uses the explicit Office preference before theme colors', () => {
    expect(resolveOfficeDarkTheme({ isDarkTheme: true, bodyBackgroundColor: '#ffffff' })).toBe(true)
    expect(resolveOfficeDarkTheme({ isDarkTheme: false, bodyBackgroundColor: '#000000' })).toBe(false)
  })

  it('uses background colors on older clients', () => {
    expect(resolveOfficeDarkTheme({ bodyBackgroundColor: '#242424' })).toBe(true)
    expect(resolveOfficeDarkTheme({ bodyBackgroundColor: '#FAFAFA' })).toBe(false)
  })

  it('leaves unavailable or invalid host theme information unresolved', () => {
    expect(resolveOfficeDarkTheme(undefined)).toBeUndefined()
    expect(resolveOfficeDarkTheme({ bodyBackgroundColor: '' })).toBeUndefined()
    expect(resolveOfficeDarkTheme({ bodyBackgroundColor: 'invalid' })).toBeUndefined()
  })

  it('uses the system preference outside Office', () => {
    vi.stubGlobal('Office', undefined)
    vi.stubGlobal('matchMedia', vi.fn().mockReturnValue({ matches: true }))
    expect(officeUsesDarkTheme()).toBe(true)
  })

  it('prefers a light Word theme to a dark system preference', () => {
    vi.stubGlobal('Office', { context: { officeTheme: { isDarkTheme: false } } })
    vi.stubGlobal('matchMedia', vi.fn().mockReturnValue({ matches: true }))
    expect(officeUsesDarkTheme()).toBe(false)
  })

  it('falls back to light when neither host nor system information is available', () => {
    vi.stubGlobal('Office', undefined)
    vi.stubGlobal('matchMedia', undefined)
    expect(officeUsesDarkTheme()).toBe(false)
  })
})
