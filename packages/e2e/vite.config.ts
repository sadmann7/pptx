import { defineConfig } from "vite";

/**
 * Serves the test harness page at / and the generated fixture decks
 * (fixtures/*.pptx) as static files at /fixtures/*.
 */
export default defineConfig({
  root: "harness",
  publicDir: "../fixtures",
  server: {
    fs: {
      // Allow serving workspace package sources (@diceui/pptx-core resolves to src/).
      allow: ["../.."],
    },
  },
});
