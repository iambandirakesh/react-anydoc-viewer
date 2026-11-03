import { readFileSync } from "node:fs";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { fileURLToPath } from "node:url";
import { resolve } from "node:path";

const here = fileURLToPath(new URL(".", import.meta.url));
const projectRoot = fileURLToPath(new URL("../../", import.meta.url));
const { version } = JSON.parse(readFileSync(resolve(projectRoot, "package.json"), "utf8")) as {
  version: string;
};

/**
 * The demo consumes the library source directly (not `dist`), so a save in
 * `src/` hot-reloads the demo. The published package is verified separately by
 * `npm run test:package`.
 */
export default defineConfig({
  root: here,
  // GitHub Pages serves the site from `/<repo>/`; the deploy workflow sets this.
  base: process.env.DEMO_BASE ?? "/",
  plugins: [react()],
  // The demo renders `src/index.ts` directly, and that module reads the version
  // the library build injects, so the demo build has to inject it too.
  define: { __RADV_VERSION__: JSON.stringify(version) },
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
