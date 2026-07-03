import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // Component tests need a DOM; the parser also requires DOMParser.
    environment: "happy-dom",
    include: ["src/**/*.test.{ts,tsx}"],
  },
});
