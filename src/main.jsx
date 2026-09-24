import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.jsx'

// No popup/opener special-casing needed here — dataverseAuth.js redirects the whole tab to
// Microsoft and back (no popup window), so the SAME tab that started the sign-in is the one that
// lands back here afterward, and ensureInitialized() (called from SignInGate) picks up the
// pending ?code= via a plain page load.
createRoot(document.getElementById('root')).render(
  <StrictMode>
    <App />
  </StrictMode>,
)
