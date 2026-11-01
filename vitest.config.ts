import { readFileSync } from "node:fs";
import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";
import { fileURLToPath } from "node:url";

const { version } = JSON.parse(
  readFileSync(fileURLToPath(new URL("./package.json", import.meta.url)), "utf8"),
) as { version: string };

export default defineConfig({
  plugins: [react()],
  // Same replacement the library build applies, so tests of `src/index.ts` see
  // the real package version instead of an undefined global.
  define: { __RADV_VERSION__: JSON.stringify(version) },
  test: {
    globals: true,
    // happy-dom is used instead of jsdom: jsdom's `html-encoding-sniffer` shim
    // requires an ESM-only module that Node < 22.12 cannot `require`.
    environment: "happy-dom",
    setupFiles: [fileURLToPath(new URL("./tests/setup.ts", import.meta.url))],
    include: ["tests/**/*.test.{ts,tsx}"],
    exclude: ["tests/e2e/**", "node_modules/**", "dist/**"],
    testTimeout: 20000,
    restoreMocks: true,
  },
});
