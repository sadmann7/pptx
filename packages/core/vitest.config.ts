import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "happy-dom",
    include: ["src/**/*.test.ts"],
    // The suite cost is dominated by per-worker environment setup and
    // module-graph imports, not the tests. Reusing workers across files
    // (no isolation) cuts both; tests build their own DOM elements and
    // don't rely on a fresh document per file.
    pool: "threads",
    isolate: false,
  },
});
