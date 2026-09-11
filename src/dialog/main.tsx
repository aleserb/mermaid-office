import { createRoot } from 'react-dom/client'
import { DialogApp } from './DialogApp'

Office.onReady(() => {
  createRoot(document.getElementById('root')!).render(<DialogApp />)
})
