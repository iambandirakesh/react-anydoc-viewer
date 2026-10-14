import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";
import { fileURLToPath } from "node:url";

export default defineConfig({
  plugins: [react()],
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
