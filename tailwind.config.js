/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./maximo_project/templates/**/*.html",
    "./maximo_app/templates/**/*.html",
    "./static/js/*.js",
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

