import { defineConfig, devices } from "@playwright/test";

/**
 * Browser integration tests.
 *
 * They run against the demo application, which builds its own sample deck in the
 * browser, so the suite needs no fixture files and no network access. The demo dev
 * server is started automatically (and an already-running one on port 5273 is
 * reused, which is what `npm run dev` gives you).
 *
 * `channel: "chrome"` uses the Chrome already installed on the machine instead of
 * downloading a bundled browser. Remove it and run
 * `npx playwright install chromium` on a CI machine with no Chrome.
 */
export default defineConfig({
  testDir: "./tests/e2e",
  fullyParallel: false,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 1 : 0,
  workers: 1,
  reporter: process.env.CI ? [["list"], ["html", { open: "never" }]] : [["list"]],
  timeout: 45_000,
  expect: { timeout: 10_000 },
  use: {
    // `localhost` (not 127.0.0.1): Vite binds the IPv6 loopback by default, and
    // Playwright must probe the same origin to reuse a running `npm run dev`.
    baseURL: "http://localhost:5273",
    trace: "retain-on-failure",
    screenshot: "only-on-failure",
    viewport: { width: 1280, height: 900 },
  },
  projects: [
    {
      name: "chromium",
      use: { ...devices["Desktop Chrome"], channel: "chrome" },
    },
  ],
  webServer: {
    command: "npm run dev -- --port 5273 --strictPort",
    url: "http://localhost:5273",
    reuseExistingServer: true,
    timeout: 60_000,
    stdout: "pipe",
  },
});
