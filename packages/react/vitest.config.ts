import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "happy-dom",
    globals: true,
    include: ["src/**/*.test.{ts,tsx}"],
    setupFiles: ["src/tests/setup.ts"],
    // The dominant suite cost is importing the package graph (react +
    // core) once per isolated worker, not the tests themselves. Threads
    // share the transform cache better than forks, and disabling
    // isolation lets workers reuse the loaded module graph across files.
    // Safe here: tests create their own store per test and don't mutate
    // module-level state.
    pool: "threads",
    isolate: false,
  },
});
