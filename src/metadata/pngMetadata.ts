import { base64ToBytes, bytesToBase64 } from './base64'
import { parseDiagramPayload, type DiagramPayload } from './payload'

const PNG_SIGNATURE = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10])
const KEYWORD = 'mermaid-office'
const textEncoder = new TextEncoder()
const textDecoder = new TextDecoder()

export function embedPayloadInPng(base64Png: string, payload: DiagramPayload): string {
  const png = base64ToBytes(base64Png)
  assertPng(png)

  const iendOffset = findChunkOffset(png, 'IEND')
  if (iendOffset === -1) {
    throw new Error('PNG does not contain an IEND chunk.')
  }

  const metadataChunk = createInternationalTextChunk(JSON.stringify(payload))
  const result = new Uint8Array(png.length + metadataChunk.length)
  result.set(png.subarray(0, iendOffset), 0)
  result.set(metadataChunk, iendOffset)
  result.set(png.subarray(iendOffset), iendOffset + metadataChunk.length)
  return bytesToBase64(result)
}

export function setPngPhysicalWidth(base64Png: string, widthPoints: number): string {
  const png = base64ToBytes(base64Png)
  assertPng(png)
  if (!Number.isFinite(widthPoints) || widthPoints <= 0) {
    throw new Error('PNG physical width must be a positive number.')
  }

  const ihdrOffset = findChunkOffset(png, 'IHDR')
  if (ihdrOffset === -1) {
    throw new Error('PNG does not contain an IHDR chunk.')
  }

  const pixelWidth = readUint32(png, ihdrOffset + 8)
  const widthMeters = (widthPoints / 72) * 0.0254
  const pixelsPerMeter = Math.max(1, Math.round(pixelWidth / widthMeters))
  const densityData = new Uint8Array(9)
  writeUint32(densityData, 0, pixelsPerMeter)
  writeUint32(densityData, 4, pixelsPerMeter)
  densityData[8] = 1
  const densityChunk = createChunk('pHYs', densityData)

  const existingOffset = findChunkOffset(png, 'pHYs')
  if (existingOffset !== -1) {
    const existingLength = readUint32(png, existingOffset) + 12
    const result = new Uint8Array(png.length - existingLength + densityChunk.length)
    result.set(png.subarray(0, existingOffset), 0)
    result.set(densityChunk, existingOffset)
    result.set(
      png.subarray(existingOffset + existingLength),
      existingOffset + densityChunk.length,
    )
    return bytesToBase64(result)
  }

  const insertionOffset = ihdrOffset + readUint32(png, ihdrOffset) + 12
  const result = new Uint8Array(png.length + densityChunk.length)
  result.set(png.subarray(0, insertionOffset), 0)
  result.set(densityChunk, insertionOffset)
  result.set(png.subarray(insertionOffset), insertionOffset + densityChunk.length)
  return bytesToBase64(result)
}

export function readPayloadFromPng(base64Png: string): DiagramPayload | null {
  const png = base64ToBytes(base64Png)
  assertPng(png)

  let offset = PNG_SIGNATURE.length
  while (offset + 12 <= png.length) {
    const length = readUint32(png, offset)
    const type = textDecoder.decode(png.subarray(offset + 4, offset + 8))
    const dataStart = offset + 8
    const dataEnd = dataStart + length

    if (dataEnd + 4 > png.length) {
      throw new Error('PNG contains a truncated chunk.')
    }

    if (type === 'iTXt') {
      const data = png.subarray(dataStart, dataEnd)
      const keywordEnd = data.indexOf(0)
      const keyword = textDecoder.decode(data.subarray(0, keywordEnd))
      if (keyword === KEYWORD) {
        const expectedCrc = readUint32(png, dataEnd)
        const actualCrc = crc32(png.subarray(offset + 4, dataEnd))
        if (actualCrc !== expectedCrc) {
          throw new Error('PNG Mermaid metadata failed its integrity check.')
        }

        const textStart = keywordEnd + 5
        return parseDiagramPayload(textDecoder.decode(data.subarray(textStart)))
      }
    }

    offset = dataEnd + 4
  }

  return null
}

function createInternationalTextChunk(text: string): Uint8Array {
  const keyword = textEncoder.encode(KEYWORD)
  const value = textEncoder.encode(text)
  const data = new Uint8Array(keyword.length + value.length + 5)
  data.set(keyword)
  data[keyword.length] = 0
  data[keyword.length + 1] = 0
  data[keyword.length + 2] = 0
  data[keyword.length + 3] = 0
  data[keyword.length + 4] = 0
  data.set(value, keyword.length + 5)

  return createChunk('iTXt', data)
}

function createChunk(typeName: string, data: Uint8Array): Uint8Array {
  const type = textEncoder.encode(typeName)
  const chunk = new Uint8Array(data.length + 12)
  writeUint32(chunk, 0, data.length)
  chunk.set(type, 4)
  chunk.set(data, 8)
  writeUint32(chunk, data.length + 8, crc32(chunk.subarray(4, data.length + 8)))
  return chunk
}

function assertPng(bytes: Uint8Array) {
  if (
    bytes.length < PNG_SIGNATURE.length ||
    !PNG_SIGNATURE.every((value, index) => bytes[index] === value)
  ) {
    throw new Error('Image data is not a valid PNG.')
  }
}

function findChunkOffset(png: Uint8Array, targetType: string): number {
  let offset = PNG_SIGNATURE.length
  while (offset + 12 <= png.length) {
    const length = readUint32(png, offset)
    const type = textDecoder.decode(png.subarray(offset + 4, offset + 8))
    if (type === targetType) {
      return offset
    }
    offset += length + 12
  }
  return -1
}

function readUint32(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(offset)
}

function writeUint32(bytes: Uint8Array, offset: number, value: number) {
  new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).setUint32(offset, value)
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff
  for (const byte of bytes) {
    crc ^= byte
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0)
    }
  }
  return (crc ^ 0xffffffff) >>> 0
}
