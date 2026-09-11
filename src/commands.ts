import { renderMermaid } from './mermaid/render'
import { insertDiagram } from './word/insertDiagram'

export const DEFAULT_DIAGRAM = `flowchart LR
    A[Start] --> B[Finish]`

async function insertDefaultDiagram(event: Office.AddinCommands.Event) {
  try {
    const svg = await renderMermaid(DEFAULT_DIAGRAM)
    await insertDiagram(svg, DEFAULT_DIAGRAM)
  } catch (error) {
    console.error('Unable to insert the default Mermaid diagram.', error)
  } finally {
    event.completed()
  }
}

Office.onReady(() => {
  Office.actions.associate('insertDefaultDiagram', insertDefaultDiagram)
})
