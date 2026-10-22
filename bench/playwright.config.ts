import { defineConfig, devices } from "@playwright/test";

/**
 * Benchmark configuration.
 *
 * Deliberately separate from `playwright.config.ts`: this suite measures how long
 * a file takes to appear, so it must not run in the ordinary browser suite (where
 * a machine under load would make it flaky) and it needs a far longer timeout.
 *
 *   npm run bench            # median load time per format
 *   BENCH_PROFILE=1 npx playwright test -c bench   # + a CPU profile of each load
 */
// BENCH_PORT runs the benchmark against a dev server of its own, which is what a
// cold-cache measurement needs (see the README's note on first-load dev cost).
const port = Number(process.env.BENCH_PORT ?? 5273);
const baseURL = `http://localhost:${port}`;

export default defineConfig({
  testDir: ".",
  fullyParallel: false,
  workers: 1,
  reporter: [["list"]],
  timeout: 300_000,
  expect: { timeout: 180_000 },
  use: {
    baseURL,
    trace: "off",
    screenshot: "off",
    viewport: { width: 1280, height: 900 },
  },
  projects: [{ name: "chromium", use: { ...devices["Desktop Chrome"], channel: "chrome" } }],
  webServer: {
    command: `npm run dev -- --port ${port} --strictPort`,
    url: baseURL,
    reuseExistingServer: true,
    timeout: 60_000,
    stdout: "pipe",
  },
});
