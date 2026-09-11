export type DiagramFormat = 'svg' | 'png'

export function isSvgInsertionSupported(): boolean {
  return (
    typeof Office !== 'undefined' &&
    Office.context.requirements.isSetSupported('ImageCoercion', '1.2')
  )
}

function setSelectedData(data: string, coercionType: Office.CoercionType): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(data, { coercionType }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve()
      } else {
        reject(new Error(result.error.message))
      }
    })
  })
}

async function svgToPngBase64(svg: string): Promise<string> {
  const svgDocument = new DOMParser().parseFromString(svg, 'image/svg+xml')
  const root = svgDocument.documentElement
  const viewBox = root.getAttribute('viewBox')?.split(/\s+/).map(Number)
  const sourceWidth = viewBox?.[2] || Number.parseFloat(root.getAttribute('width') || '') || 1200
  const sourceHeight = viewBox?.[3] || Number.parseFloat(root.getAttribute('height') || '') || 800
  const scale = Math.min(2, 4096 / Math.max(sourceWidth, sourceHeight))
  const width = Math.max(1, Math.round(sourceWidth * scale))
  const height = Math.max(1, Math.round(sourceHeight * scale))
  const url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }))

  try {
    const image = new Image()
    image.src = url
    await image.decode()

    const canvas = window.document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    const context = canvas.getContext('2d')
    if (!context) {
      throw new Error('This browser cannot create the PNG fallback.')
    }

    context.drawImage(image, 0, 0, width, height)
    return canvas.toDataURL('image/png').split(',', 2)[1]
  } finally {
    URL.revokeObjectURL(url)
  }
}

export async function insertDiagram(svg: string): Promise<DiagramFormat> {
  if (typeof Office === 'undefined' || !Office.context?.document) {
    throw new Error('Open Mermaid Office inside Microsoft Word to insert a diagram.')
  }

  if (isSvgInsertionSupported()) {
    await setSelectedData(svg, Office.CoercionType.XmlSvg)
    return 'svg'
  }

  const png = await svgToPngBase64(svg)
  await setSelectedData(png, Office.CoercionType.Image)
  return 'png'
}
