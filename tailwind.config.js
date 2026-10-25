/** @type {import('tailwindcss').Config} */
// Colors are the corporate Synapse/HMIS Design System tokens (--ds-* in src/index.css, from the
// "Code App Design System — Guidelines" the user shared). Since 2026-10-25 the app uses the
// system's PRIMARY lane — teal #126e5c — for actions, navigation, active states and focus, per the
// user ("no gold — that's why I sent you the design"); it replaced the gold Analytics-lane accent
// chosen 2026-09-27. ink/muted are the system's neutral-900/neutral-500. Only the tokens are ported:
// the design system's React package pins React 18, this app runs React 19.
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        primary: {
          DEFAULT: '#126e5c', // --ds-color-primary (600)
          strong: '#0d594b', // 700 — hover / text on light tints
          soft: '#d9f8ec', // --ds-color-primary-soft
          deep: '#023635', // --ds-color-primary-deep
          mint: '#40ffb8', // --ds-color-primary-mint (accent highlight)
          50: '#f0faf7',
          100: '#d9f3ea',
          200: '#b7e7d7',
          300: '#82d2ba',
          400: '#4bb69a',
          500: '#278f76',
          600: '#126e5c',
          700: '#0d594b',
          800: '#0a463c',
          900: '#073a33',
          950: '#03251f',
        },
        ink: '#0f172a',
        muted: '#64748b',
      },
      fontFamily: {
        sans: ['Urbanist', 'ui-sans-serif', 'system-ui', '-apple-system', 'BlinkMacSystemFont', 'Segoe UI', 'Arial', 'sans-serif'],
      },
      borderRadius: {
        control: '8px', // --ds-radius-control (buttons, inputs)
      },
      boxShadow: {
        // --ds-shadow-button / -button-hover, --ds-shadow-sm / -lg, --ds-shadow-dialog
        button: '0 10px 22px rgba(2, 54, 53, 0.17), inset 0 1px 0 rgba(255, 255, 255, 0.13)',
        'button-hover': '0 13px 28px rgba(2, 54, 53, 0.23)',
        'ds-sm': '0 13px 28px rgba(82, 82, 100, 0.11), inset 0 1px 0 rgba(255, 255, 255, 0.88)',
        'ds-lg': '0 30px 70px rgba(76, 73, 105, 0.19), inset 0 1px 0 rgba(255, 255, 255, 0.9)',
        dialog: '0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1)',
      },
      transitionTimingFunction: {
        settle: 'cubic-bezier(.16,1,.3,1)',
        enter: 'cubic-bezier(.2,.7,.2,1)',
        exit: 'cubic-bezier(.4,0,.7,.2)',
      },
    },
  },
  plugins: [],
}
