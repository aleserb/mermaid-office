import { base64ToBytes, base64ToText } from './base64'
import type { DiagramPayload } from './payload'
import { readPayloadFromPng } from './pngMetadata'
import { readPayloadFromSvg } from './svgMetadata'

const PNG_SIGNATURE = [137, 80, 78, 71, 13, 10, 26, 10]

function stripDataUrl(value: string): string {
  const separator = value.indexOf(',')
  return value.startsWith('data:') && separator >= 0
    ? value.slice(separator + 1)
    : value
}

export function readPayloadFromImage(base64Image: string): DiagramPayload | null {
  const base64 = stripDataUrl(base64Image)
  const bytes = base64ToBytes(base64)

  if (
    bytes.length >= PNG_SIGNATURE.length &&
    PNG_SIGNATURE.every((value, index) => bytes[index] === value)
  ) {
    return readPayloadFromPng(base64)
  }

  const text = base64ToText(base64).trim()
  return /^<svg[\s>]/i.test(text) ? readPayloadFromSvg(text) : null
}
