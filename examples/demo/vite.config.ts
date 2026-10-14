import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { fileURLToPath } from "node:url";

const here = fileURLToPath(new URL(".", import.meta.url));
const projectRoot = fileURLToPath(new URL("../../", import.meta.url));

/**
 * The demo consumes the library source directly (not `dist`), so a save in
 * `src/` hot-reloads the demo. The published package is verified separately by
 * `npm run test:package`.
 */
export default defineConfig({
  root: here,
  plugins: [react()],
  resolve: {
    // Exact matches only: the bare package name must not swallow the CSS subpath.
    alias: [
      {
        find: /^react-anydoc-viewer\/styles\.css$/,
        replacement: fileURLToPath(new URL("../../src/styles.css", import.meta.url)),
      },
      {
        find: /^react-anydoc-viewer$/,
        replacement: fileURLToPath(new URL("../../src/index.ts", import.meta.url)),
      },
    ],
  },
  server: {
    port: 5273,
    fs: {
      allow: [projectRoot],
    },
  },
  build: {
    outDir: "dist",
    emptyOutDir: true,
  },
});
