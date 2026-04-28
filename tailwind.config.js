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
          50: '#FFFFFF',
          100: '#FFFFFF',
          200: '#F2EBDC',
          300: '#EFE7DA',
          400: '#B8AC95',
          500: '#A89B82',
        },
      },
      fontFamily: {
        sans: ['"Pretendard Variable"', 'Pretendard', 'ui-sans-serif', 'system-ui', '-apple-system', 'sans-serif'],
        display: ['"Pretendard Variable"', 'Pretendard', 'sans-serif'],
      },
      fontSize: {
        xs: ['13px', { lineHeight: '1.5' }],
        sm: ['15px', { lineHeight: '1.5' }],
        base: ['17px', { lineHeight: '1.55' }],
        lg: ['19px', { lineHeight: '1.5' }],
        xl: ['21px', { lineHeight: '1.45' }],
        '2xl': ['25px', { lineHeight: '1.35' }],
        '3xl': ['31px', { lineHeight: '1.25' }],
        '4xl': ['37px', { lineHeight: '1.2' }],
      },
    },
  },
  plugins: [],
}
