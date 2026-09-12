export function SyntaxHelpLink() {
  return (
    <a
      className="syntax-help"
      href="https://mermaid.js.org/intro/syntax-reference.html"
      target="_blank"
      rel="noreferrer"
      aria-label="Mermaid syntax"
      title="Mermaid syntax"
    >
      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" aria-hidden="true" focusable="false">
        <path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9l-6-6Z"
          stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" />
        <path d="M14 3v6h6M11 14h1v4m-1 0h2"
          stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
        <circle cx="12" cy="11.5" r="1" fill="currentColor" />
      </svg>
    </a>
  )
}
