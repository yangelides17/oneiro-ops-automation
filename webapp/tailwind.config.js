/** @type {import('tailwindcss').Config} */
export default {
  content: [
    './index.html',
    './src/**/*.{js,jsx,ts,tsx}'
  ],
  theme: {
    extend: {
      colors: {
        navy: {
          50:  '#eef1f7',
          100: '#cdd5e6',
          200: '#9cadc8',
          300: '#6b85aa',
          400: '#3a5d8c',
          DEFAULT: '#1B2A4A',
          700: '#162038',
          800: '#111827',
          900: '#0d1520'
        },
        gold: {
          light: '#f0d878',
          DEFAULT: '#C9A227',
          dark:  '#a0821f'
        }
      },
      fontFamily: {
        sans: [
          '-apple-system', 'BlinkMacSystemFont', '"Segoe UI"',
          'Roboto', 'Helvetica', 'Arial', 'sans-serif'
        ]
      }
    }
  },
  plugins: []
}
