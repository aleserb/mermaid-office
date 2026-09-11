# Mermaid Office

Mermaid Office is a client-side Microsoft Word add-in for creating and editing
Mermaid diagrams. Mermaid renders locally in the Office task pane; no
application server receives diagram source or document content.

See [design.md](design.md) for the product goals, user experience, architecture,
compatibility strategy, and delivery plan.

## Technology

- React, TypeScript, and Vite
- Office.js with SVG insertion through ImageCoercion 1.2
- Mermaid.js and DOMPurify
- CodeMirror 6
- Fluent UI React v9
- Vitest and Testing Library
- GitHub Pages

SVG is the primary format on supported Word desktop clients. The add-in creates
a PNG locally when SVG insertion is unavailable, including Word on the web.

## Development

```bash
npm install
npm run dev
```

The regular browser view supports editor and preview development. Diagram
insertion must be tested by sideloading `manifest.xml` in Microsoft Word.
After manifest changes, remove and upload the manifest again because Word caches
the sideloaded manifest separately from the hosted web application.

When the task pane is open, selecting an inserted Mermaid diagram loads its
stored source into the editor automatically. Saving replaces the selected
picture in place while preserving its displayed width and alternative text.
Use **New diagram** to leave edit mode.

Use **Expand editor** for a larger two-column editing dialog with a live
preview. Parse failures are shown in the editor gutter, keep the last valid
preview visible, and prevent saving until the source is valid.
Preview controls support zooming in and out, resetting to fit, and dragging a
zoomed diagram to pan. Hold Ctrl (Windows) or Command (macOS) while scrolling
over the preview to zoom without taking over normal page scrolling.

The ribbon's **Open editor** command opens this dialog directly without opening
the task pane. Because this changes `manifest.xml`, remove and upload the
manifest again after deployment.

## Checks

```bash
npm run lint
npm test
npm run build
```

## Deployment

Run `npm run deploy` to build the application and publish `dist` to the
`gh-pages` branch. Configure the repository's Pages source as the root of that
branch before the first deployment.

GitHub Pages for a private repository requires a GitHub plan that supports
private Pages sites. Otherwise, make the repository public or upgrade the
account before enabling Pages.
