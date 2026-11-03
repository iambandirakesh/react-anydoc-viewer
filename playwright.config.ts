import { defineConfig, devices } from "@playwright/test";

/**
 * Browser integration tests.
 *
 * They run against the demo application, which builds its own sample deck in the
 * browser, so the suite needs no fixture files and no network access. The demo dev
 * server is started automatically (and an already-running one on port 5273 is
 * reused, which is what `npm run dev` gives you).
 *
 * Locally, `channel: "chrome"` uses the Chrome already installed on the machine
 * instead of downloading a bundled browser. In CI (`CI` is set) the bundled
 * Chromium is used, so runs are reproducible; the CI workflow installs it with
 * `npx playwright install --with-deps chromium`.
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
      use: { ...devices["Desktop Chrome"], ...(process.env.CI ? {} : { channel: "chrome" }) },
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
