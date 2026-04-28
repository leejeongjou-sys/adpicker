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
        sans: ['"Pretendard Variable"', 'Pretendard', 'ui-sans-serif', 'system-ui', '-apple-system', 'sans-serif'],
        display: ['"Pretendard Variable"', 'Pretendard', 'sans-serif'],
      },
    },
  },
  plugins: [],
}
