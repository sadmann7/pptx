import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "happy-dom",
    globals: true,
    include: ["src/**/*.test.{ts,tsx}"],
    setupFiles: ["src/tests/setup.ts"],
    // Importing react and core per isolated worker dominates the suite's cost.
    // Sharing the module graph is safe because tests create their own store and
    // don't mutate module state.
    pool: "threads",
    isolate: false,
  },
});
