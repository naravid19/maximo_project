/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./maximo_app/templates/**/*.html",
    "./node_modules/flowbite/**/*.js"
],
  theme: {
    extend: {
      fontFamily: {
        sans: ['"Bai Jamjuree"', 'sans-serif'],
      },
    },
  },
  plugins: [
    require('flowbite/plugin'),
  ],
};

