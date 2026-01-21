import type { Config } from "tailwindcss";

const config: Config = {
  content: ["./src/app/**/*.{ts,tsx}", "./src/components/**/*.{ts,tsx}"],
  theme: {
    extend: {
      colors: {
        ink: "#1f2430",
        coral: "#f36b5d",
        sand: "#f6efe8",
        moss: "#4d6a5a",
        sky: "#d6e3f3"
      }
    }
  },
  plugins: []
};

export default config;
