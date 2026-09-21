/** @type {import('tailwindcss').Config} */
// Ported from the Segmentation project's tailwind.config.js — same "Andalusia Group" brand
// palette (gold/ink/muted) so this tool looks like a sibling of the existing dashboard rather
// than a one-off design.
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        gold: {
          DEFAULT: '#AE8C67',
          strong: '#9a7b57',
        },
        ink: '#1f2430',
        muted: '#6a7380',
      },
    },
  },
  plugins: [],
}
