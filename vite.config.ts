import { readFileSync } from "node:fs";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { fileURLToPath } from "node:url";
import { resolve } from "node:path";

const root = fileURLToPath(new URL(".", import.meta.url));
const { version } = JSON.parse(readFileSync(resolve(root, "package.json"), "utf8")) as {
  version: string;
};

export default defineConfig({
  plugins: [react()],
  // The published `version` export must match the package it ships in, so the
  // number is replaced at build time instead of being written down twice.
  define: { __RADV_VERSION__: JSON.stringify(version) },
  build: {
    target: "es2019",
    // Source maps are shipped so consumers can debug the (minified) output.
    sourcemap: true,
    minify: true,
    cssMinify: true,
    lib: {
      entry: resolve(root, "src/index.ts"),
      name: "ReactAnydocViewer",
      formats: ["es", "cjs", "umd"],
      fileName: (format) =>
        format === "es" ? "index.js" : format === "cjs" ? "index.cjs" : "index.umd.cjs",
      cssFileName: "styles",
    },
    rollupOptions: {
      // `pdfjs-dist` is external so the ~1MB parser stays out of the bundle and
      // is only ever fetched when a PDF is opened (see `src/parser/pdf.ts`).
      // Both builds are listed: the modern one and the legacy one that older
      // engines need.
      external: [
        "react",
        "react-dom",
        "react/jsx-runtime",
        "react-dom/client",
        "pdfjs-dist",
        "pdfjs-dist/build/pdf.worker.min.mjs",
        "pdfjs-dist/legacy/build/pdf.mjs",
        "pdfjs-dist/legacy/build/pdf.worker.min.mjs",
      ],
      output: {
        globals: {
          react: "React",
          "react-dom": "ReactDOM",
          "react/jsx-runtime": "ReactJSXRuntime",
          "react-dom/client": "ReactDOMClient",
        },
        exports: "named",
      },
    },
  },
});
