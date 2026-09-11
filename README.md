# Mermaid Office

Mermaid Office is a client-side Microsoft Word add-in for creating and editing
Mermaid diagrams. Mermaid renders locally in the Office task pane; no
application server receives diagram source or document content.

## Technology

- React, TypeScript, and Vite
- Office.js with SVG insertion through ImageCoercion 1.2
- Mermaid.js and DOMPurify
- CodeMirror 6
- Fluent UI React v9
- Vitest and Testing Library
- GitHub Pages and GitHub Actions

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

Pushes to `main` build and deploy the static application with
`.github/workflows/deploy-pages.yml`. Configure the repository's Pages source
as **GitHub Actions** before the first deployment.
