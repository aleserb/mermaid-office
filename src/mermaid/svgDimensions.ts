export function getSvgDimensions(root: Element): { width: number; height: number } {
  const viewBox = root.getAttribute('viewBox')?.split(/\s+/).map(Number)
  return {
    width: viewBox?.[2] || Number.parseFloat(root.getAttribute('width') || '') || 1200,
    height: viewBox?.[3] || Number.parseFloat(root.getAttribute('height') || '') || 800,
  }
}
