import React from 'react'
import ReactDOM from 'react-dom/client'
import { BrowserRouter } from 'react-router-dom'
import App from './App'
import './index.css'

// Reset any persisted page zoom on every load. iOS Safari remembers the
// pinch-zoom level across reloads, so a user who accidentally zooms in the
// field gets "stuck" — even a refresh keeps the old zoom. Momentarily
// clamping the viewport to maximum-scale=1 snaps the scale back to 1x;
// we then restore the zoom-enabled viewport so the user can still pinch
// to zoom (e.g. to read a dense table) whenever they want.
function resetPageZoom() {
  const meta = document.querySelector('meta[name="viewport"]')
  if (!meta) return
  const zoomable = 'width=device-width, initial-scale=1.0'
  meta.setAttribute('content', zoomable + ', maximum-scale=1.0')
  // Restore after the clamp has applied so pinch-zoom stays available.
  setTimeout(() => meta.setAttribute('content', zoomable), 350)
}
// pageshow fires on the initial load AND on reloads / back-forward and
// bfcache restores — covering every way the user lands back on the page.
window.addEventListener('pageshow', resetPageZoom)
resetPageZoom()

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <BrowserRouter>
      <App />
    </BrowserRouter>
  </React.StrictMode>
)
