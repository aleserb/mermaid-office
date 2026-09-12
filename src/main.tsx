import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.tsx'
import { PaneApp } from './pane/PaneApp'

if (typeof Office !== 'undefined') {
  Office.onReady()
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    {new URLSearchParams(window.location.search).get('view') === 'pane' ? <PaneApp /> : <App />}
  </StrictMode>,
)
