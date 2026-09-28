/** @type {import('tailwindcss').Config} */
// 2026-09-27: colors realigned to the real corporate "Synapse/HMIS Design System" tokens (found
// at hmis-design-system/src/design-system/tokens.css) rather than our own hand-picked gold —
// specifically its "Analytics lane" scale, which is that design system's OWN designated color for
// reporting/analytics tools exactly like this one (the system-wide default is a teal, reserved for
// non-analytics product surfaces). Values below are that lane's 500/600 shades verbatim
// (--ds-color-analytics-500/600). ink/muted are that same design system's neutral-900/neutral-500
// — close to, but not identical to, the values this project already had (validates the original
// choice rather than overturning it). Token NAMES (gold/ink/muted) are kept as-is rather than
// renamed throughout every component. NOT installed as a live npm dependency: the design system
// package pins React 18.3.1 against this project's React 19, so only its CSS tokens are ported,
// not its component code (see DECISIONS-equivalent discussion — no live coupling, no dual-React
// risk).
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        gold: {
          DEFAULT: '#c99b3b',
          strong: '#a07118',
          soft: '#fbefd6',
        },
        ink: '#0f172a',
        muted: '#64748b',
      },
      fontFamily: {
        sans: ['Urbanist', 'ui-sans-serif', 'system-ui', '-apple-system', 'BlinkMacSystemFont', 'Segoe UI', 'Arial', 'sans-serif'],
      },
      boxShadow: {
        'gold-button': '0 10px 22px rgba(160, 113, 24, 0.17), inset 0 1px 0 rgba(255, 255, 255, 0.13)',
        'gold-button-hover': '0 13px 28px rgba(160, 113, 24, 0.23)',
      },
    },
  },
  plugins: [],
}
