import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "happy-dom",
    globals: true,
    include: ["src/**/*.test.{ts,tsx}"],
    setupFiles: ["src/tests/setup.ts"],
    // Most of the suite's cost is importing react and core once per isolated worker.
    // Threads share the transform cache better than forks, and without isolation
    // workers reuse the module graph across files, which is safe because tests
    // create their own store and don't mutate module state.
    pool: "threads",
    isolate: false,
  },
});
