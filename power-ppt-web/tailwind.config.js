/** @type {import('tailwindcss').Config} */
// Salasar brand tokens. Colours and gradient are the authorised set only —
// Blue #1A3A8F, Green #7AC143, Mid Blue #0070C0, Slate #4D4D4D (body), White.
export default {
  content: ["./index.html", "./src/**/*.{ts,tsx}"],
  theme: {
    extend: {
      colors: {
        brand: {
          blue: "#1A3A8F",
          green: "#7AC143",
          midblue: "#0070C0",
          slate: "#4D4D4D",
        },
      },
      fontFamily: {
        sans: ["Poppins", "system-ui", "sans-serif"],
      },
      backgroundImage: {
        // Full-bleed brand gradient (white text only, never behind body copy).
        "brand-gradient":
          "linear-gradient(135deg,#7AC143 0%,#0070C0 50%,#1A3A8F 100%)",
        // Footer bar gradient.
        "brand-bar": "linear-gradient(90deg,#7AC143,#1A3A8F)",
      },
      borderRadius: {
        card: "0.75rem",
      },
    },
  },
  plugins: [],
};
