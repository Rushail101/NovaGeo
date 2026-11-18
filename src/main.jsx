import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import { AuthGate, PasswordRecovery, AUTH_CSS } from './auth.jsx'
import { Toaster } from './ui.jsx'
import { FONTS, CSS, EXTRA_CSS, MODULE_CSS } from './styles.js'

try { document.documentElement.dataset.theme = localStorage.getItem('fin-theme') || 'dark' } catch {}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <style>{FONTS}{CSS}{EXTRA_CSS}{MODULE_CSS}{AUTH_CSS}</style>
    <AuthGate>
      <App />
    </AuthGate>
    <PasswordRecovery />
    <Toaster />
  </React.StrictMode>,
)
