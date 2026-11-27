import { defineConfig } from "vite";

// Relative base so the built page works at the site root and at
// https://cellproof.github.io/xlsx-parity/ without rewriting asset URLs.
export default defineConfig({
  base: "./",
  build: {
    outDir: "dist",
    emptyOutDir: true,
    target: "es2022",
    // SheetJS CE is one parser. The warning is the size of that file, not a split we forgot.
    chunkSizeWarningLimit: 800,
  },
  optimizeDeps: {
    exclude: ["@ironcalc/wasm"],
  },
});
