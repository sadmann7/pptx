import { defineConfig } from "vite";

/**
 * Serves the test harness page at / and the generated fixture decks
 * (fixtures/*.pptx) as static files at the server root (e.g. /basic.pptx).
 */
export default defineConfig({
  root: "harness",
  publicDir: "../fixtures",
  server: {
    fs: {
      // Paths are relative to the "harness" root. ".." covers this package
      // (incl. node_modules/.vite prebundled deps); "../../core" covers the
      // @diceui/pptx-core sources the workspace alias resolves to.
      allow: ["..", "../../core"],
    },
  },
});
