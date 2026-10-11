import { Component } from "react";

// This app had NO error boundary anywhere — any uncaught render-time error (a bad prop, a stray
// typo like a mismatched prop name) unmounts the ENTIRE React tree, leaving a silent, blank white
// page with no on-screen indication anything went wrong at all (caught live, 2026-10-06: a
// RibbonButton/Popover prop mismatch crashed the whole app the moment any ribbon button was
// clicked — "بيودوني على صفحة فاضية" — and even the browser's own DevTools console didn't
// obviously surface it; only Vite's own terminal-side HMR error log did). This doesn't fix the
// underlying bug class (a crash is still a crash) — it just makes "something broke" visible and
// recoverable instead of indistinguishable from the app simply not loading.
export default class ErrorBoundary extends Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }

  static getDerivedStateFromError(error) {
    return { error };
  }

  componentDidCatch(error, info) {
    // eslint-disable-next-line no-console
    console.error("Uncaught render error:", error, info?.componentStack);
  }

  render() {
    if (this.state.error) {
      return (
        <div className="flex min-h-screen flex-col items-center justify-center gap-4 bg-white p-8 text-center">
          <p className="text-lg font-semibold text-ink">Something went wrong.</p>
          <p className="max-w-md text-sm text-muted">{this.state.error.message || String(this.state.error)}</p>
          <button
            type="button"
            onClick={() => window.location.reload()}
            className="btn-primary"
          >
            Reload
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}
