/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{ts,tsx}"],
  theme: {
    extend: {
      colors: {
        ink: "#172033",
        lagoon: "#5f55d6",
        ember: "#b0003a",
        dexyBlue: "#1298e8",
        dexyPurple: "#5f55d6",
        dexyRed: "#b0003a",
        cloud: "#f6f8fb",
      },
      boxShadow: {
        soft: "0 18px 50px rgba(23, 32, 51, 0.10)",
      },
    },
  },
  plugins: [],
};
