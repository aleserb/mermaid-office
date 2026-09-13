import { useEffect, useState } from 'react'

export function resolveOfficeDarkTheme(
  theme: Partial<Pick<Office.OfficeTheme, 'isDarkTheme' | 'bodyBackgroundColor'>> | undefined,
): boolean | undefined {
  if (typeof theme?.isDarkTheme === 'boolean') {
    return theme.isDarkTheme
  }
  const hex = /^#([0-9a-f]{6})$/i.exec(theme?.bodyBackgroundColor ?? '')?.[1]
  if (!hex) {
    return undefined
  }
  const red = Number.parseInt(hex.slice(0, 2), 16)
  const green = Number.parseInt(hex.slice(2, 4), 16)
  const blue = Number.parseInt(hex.slice(4, 6), 16)
  return 0.2126 * red + 0.7152 * green + 0.0722 * blue < 128
}

export function systemUsesDarkTheme(): boolean {
  return window.matchMedia?.('(prefers-color-scheme: dark)').matches ?? false
}

export function officeUsesDarkTheme(): boolean {
  const theme = typeof Office === 'undefined' ? undefined : Office.context?.officeTheme
  return resolveOfficeDarkTheme(theme) ?? systemUsesDarkTheme()
}

export function useOfficeDarkTheme(): boolean {
  const [dark, setDark] = useState(officeUsesDarkTheme)

  useEffect(() => {
    let active = true
    const refresh = () => {
      if (active) setDark(officeUsesDarkTheme())
    }
    const media = window.matchMedia?.('(prefers-color-scheme: dark)')
    media?.addEventListener('change', refresh)
    window.addEventListener('focus', refresh)
    if (typeof Office !== 'undefined') {
      void Office.onReady().then(refresh)
    }
    return () => {
      active = false
      media?.removeEventListener('change', refresh)
      window.removeEventListener('focus', refresh)
    }
  }, [])

  return dark
}
