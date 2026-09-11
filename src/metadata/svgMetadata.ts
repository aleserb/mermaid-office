import { base64ToText, textToBase64 } from './base64'
import { parseDiagramPayload, type DiagramPayload } from './payload'

const METADATA_ID = 'mermaid-office'

export function embedPayloadInSvg(svg: string, payload: DiagramPayload): string {
  const document = new DOMParser().parseFromString(svg, 'image/svg+xml')
  const parserError = document.querySelector('parsererror')
  if (parserError) {
    throw new Error('Rendered SVG is not valid XML.')
  }

  document.querySelector(`metadata#${METADATA_ID}`)?.remove()
  const metadata = document.createElementNS('http://www.w3.org/2000/svg', 'metadata')
  metadata.setAttribute('id', METADATA_ID)
  metadata.textContent = textToBase64(JSON.stringify(payload))
  document.documentElement.prepend(metadata)

  return new XMLSerializer().serializeToString(document.documentElement)
}

export function readPayloadFromSvg(svg: string): DiagramPayload | null {
  const document = new DOMParser().parseFromString(svg, 'image/svg+xml')
  const encoded = document.querySelector(`metadata#${METADATA_ID}`)?.textContent
  return encoded ? parseDiagramPayload(base64ToText(encoded)) : null
}
