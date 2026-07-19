import { defineConfig, devices } from "@playwright/test";

const PORT = 5000;

export default defineConfig({
  testDir: "./specs",
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  reporter: process.env.CI ? [["html"], ["github"]] : [["list"]],
  use: {
    baseURL: `http://localhost:${PORT}`,
    trace: "on-first-retry",
  },
  expect: {
    toHaveScreenshot: {
      // Canvas/font antialiasing differs slightly across GPUs and OSes.
      maxDiffPixelRatio: 0.02,
    },
  },
  // Baselines are named per-browser but not per-OS: ground truth is defined by
  // whichever machine generates them (see snapshotPathTemplate without {platform}).
  snapshotPathTemplate: "{testDir}/__screenshots__/{testFilePath}/{arg}-{projectName}{ext}",
  projects: [
    { name: "chromium", use: { ...devices["Desktop Chrome"] } },
    { name: "firefox", use: { ...devices["Desktop Firefox"] } },
  ],
  webServer: {
    command: `pnpm exec vite --port ${PORT} --strictPort`,
    url: `http://localhost:${PORT}`,
    reuseExistingServer: !process.env.CI,
    timeout: 30_000,
  },
});
