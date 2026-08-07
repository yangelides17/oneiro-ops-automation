import { Component } from 'react'

// Catches render/lifecycle errors from the page tree below it.
//
// Without one, a single unhandled throw anywhere in a page unmounts the
// entire React tree: the browser is left on a blank white screen with the
// route still in the URL and nothing to click, and the only way out is a
// manual reload. That happened for real — a pdf.js call against a
// destroyed worker threw synchronously out of an effect and took the whole
// app down with it. The underlying bug is fixed, but the failure mode it
// exposed is worth closing off permanently: an admin halfway through
// filing paperwork should never lose the screen with no way back.
//
// Deliberately NOT a silent recovery. The error is shown and logged, since
// swallowing it would leave the app looking fine while quietly broken.
export default class ErrorBoundary extends Component {
  constructor(props) {
    super(props)
    this.state = { error: null }
  }

  static getDerivedStateFromError(error) {
    return { error }
  }

  componentDidCatch(error, info) {
    // Keep the stack in the console — it's what makes these diagnosable
    // from a screenshot of devtools.
    console.error('Unhandled UI error:', error, info?.componentStack)
  }

  render() {
    if (!this.state.error) return this.props.children
    return (
      <div className="max-w-lg mx-auto px-4 py-16 text-center space-y-4">
        <div className="text-4xl">⚠️</div>
        <h1 className="text-lg font-black text-navy">Something went wrong on this page</h1>
        <p className="text-sm text-slate-600">
          The rest of the app is still running — use the tabs above to move on, or
          reload to come back to this page.
        </p>
        <p className="text-[11px] font-mono text-slate-400 break-words">
          {String(this.state.error?.message || this.state.error)}
        </p>
        <div className="flex items-center justify-center gap-2 pt-1">
          <button
            type="button"
            onClick={() => this.setState({ error: null })}
            className="text-xs font-bold px-3 py-2 rounded-lg bg-slate-100 text-slate-700 hover:bg-slate-200"
          >
            Try again
          </button>
          <button
            type="button"
            onClick={() => window.location.reload()}
            className="text-xs font-bold px-4 py-2 rounded-lg bg-navy text-white hover:opacity-90"
          >
            Reload page
          </button>
        </div>
      </div>
    )
  }
}
