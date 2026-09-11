import type { DiagramSize } from '../metadata/payload'

const sizes: Array<{ value: DiagramSize; label: string }> = [
  { value: 'small', label: 'Small' },
  { value: 'medium', label: 'Medium' },
  { value: 'large', label: 'Large' },
  { value: 'page-width', label: '6.5 in wide' },
]

interface SizePickerProps {
  value: DiagramSize
  onChange: (size: DiagramSize) => void
}

export function SizePicker({ value, onChange }: SizePickerProps) {
  return (
    <label className="theme-picker">
      <span>Size</span>
      <select
        aria-label="Diagram size"
        value={value}
        onChange={(event) => onChange(event.target.value as DiagramSize)}
      >
        {sizes.map((size) => (
          <option key={size.value} value={size.value}>
            {size.label}
          </option>
        ))}
      </select>
    </label>
  )
}
