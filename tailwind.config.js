/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        cream: {
          50: '#F5EFE6',
          100: '#EFE7DA',
          200: '#E8DFD0',
          300: '#D9CDB9',
          400: '#B8AC95',
          500: '#A89B82',
        },
      },
      fontFamily: {
        sans: ['"Noto Sans KR"', 'ui-sans-serif', 'system-ui', 'sans-serif'],
        serif: ['"Playfair Display"', '"Noto Serif KR"', 'ui-serif', 'Georgia', 'serif'],
        display: ['"Playfair Display"', '"Noto Serif KR"', 'serif'],
      },
    },
  },
  plugins: [],
}
