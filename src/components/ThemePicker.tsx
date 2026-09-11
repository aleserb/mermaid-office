import type { DiagramTheme } from '../metadata/payload'
import './ThemePicker.css'

const themes: Array<{ value: DiagramTheme; label: string }> = [
  { value: 'default', label: 'Default' },
  { value: 'neutral', label: 'Neutral' },
  { value: 'dark', label: 'Dark' },
  { value: 'forest', label: 'Forest' },
  { value: 'redux', label: 'Redux' },
]

interface ThemePickerProps {
  value: DiagramTheme
  onChange: (theme: DiagramTheme) => void
}

export function ThemePicker({ value, onChange }: ThemePickerProps) {
  return (
    <label className="theme-picker">
      <span>Theme</span>
      <select
        aria-label="Diagram theme"
        value={value}
        onChange={(event) => onChange(event.target.value as DiagramTheme)}
      >
        {themes.map((theme) => (
          <option key={theme.value} value={theme.value}>
            {theme.label}
          </option>
        ))}
      </select>
    </label>
  )
}
