import { defineConfig, devices } from "@playwright/test";

import { HARNESS_ORIGIN, HARNESS_PORT } from "./lib/constants";

export default defineConfig({
  testDir: "./specs",
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  reporter: process.env.CI ? [["html"], ["github"]] : [["list"]],
  use: {
    baseURL: HARNESS_ORIGIN,
    trace: "on-first-retry",
  },
  expect: {
    toHaveScreenshot: {
      // Canvas/font antialiasing differs slightly across GPUs and OSes.
      maxDiffPixelRatio: 0.02,
    },
  },
  // Baselines are named per-browser but not per-OS: whichever machine records
  // them defines them (see snapshotPathTemplate without {platform}).
  snapshotPathTemplate: "{testDir}/screenshots/{testFilePath}/{arg}-{projectName}{ext}",
  projects: [
    { name: "chromium", use: { ...devices["Desktop Chrome"] } },
    { name: "firefox", use: { ...devices["Desktop Firefox"] } },
  ],
  webServer: {
    // `pnpm exec` puts vite in its own process group, so Playwright's shutdown
    // misses it and waits forever on the orphaned server.
    command: `node node_modules/vite/bin/vite.js --port ${HARNESS_PORT} --strictPort`,
    url: HARNESS_ORIGIN,
    reuseExistingServer: !process.env.CI,
    timeout: 30_000,
  },
});
