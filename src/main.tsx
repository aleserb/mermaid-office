import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import { PaneApp } from './pane/PaneApp'
import { SettingsWindow } from './settings/SettingsWindow'

const App = new URLSearchParams(window.location.search).get('view') === 'settings' ? SettingsWindow : PaneApp

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
)
