import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // Component tests need a DOM; the parser also requires DOMParser.
    environment: "happy-dom",
    // Required for @testing-library/react auto-cleanup between tests.
    globals: true,
    include: ["src/**/*.test.{ts,tsx}"],
  },
});
