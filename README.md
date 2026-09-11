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
