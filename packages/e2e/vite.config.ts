import { createReadStream, existsSync } from "node:fs";
import { join } from "node:path";
import { defineConfig, type Plugin } from "vite";

const DECKS_DIR = join(import.meta.dirname, "decks");

/**
 * Serves ad-hoc decks from the gitignored decks/ directory at /decks/*.
 *
 * Reproducing a bug usually starts with a real deck that cannot be committed;
 * dropping it in decks/ makes it loadable as ?file=decks/whatever.pptx without
 * it ever looking like a fixture.
 */
function serveLocalDecks(): Plugin {
  return {
    name: "e2e-serve-local-decks",
    apply: "serve",
    configureServer(server) {
      server.middlewares.use("/decks", (req, res, next) => {
        const name = decodeURIComponent((req.url ?? "").split("?")[0]).replace(/^\//, "");
        const path = join(DECKS_DIR, name);
        if (!name || name.includes("..") || !existsSync(path)) return next();
        res.setHeader(
          "Content-Type",
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        );
        createReadStream(path).pipe(res);
      });
    },
  };
}

/**
 * Serves the test harness page at / and the generated fixture decks
 * (fixtures/*.pptx) as static files at the server root (e.g. /basic.pptx).
 */
export default defineConfig({
  root: "harness",
  publicDir: "../fixtures",
  plugins: [serveLocalDecks()],
  server: {
    fs: {
      // Paths are relative to the "harness" root. ".." covers this package
      // (incl. node_modules/.vite prebundled deps); "../../core" covers the
      // @diceui/pptx-core sources the workspace alias resolves to.
      allow: ["..", "../../core"],
    },
  },
});
