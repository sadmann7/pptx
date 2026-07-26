import { createReadStream, existsSync } from "node:fs";
import { resolve, sep } from "node:path";
import { defineConfig, type Plugin } from "vite";

const DECKS_DIR = resolve(import.meta.dirname, "decks");

/**
 * Resolves a request path to a file inside decks/, or undefined.
 *
 * The containment check is what makes this safe, rather than any property of
 * how the request path and the base directory combine: whatever the request
 * asks for, the file that gets served has to sit under decks/.
 */
function resolveDeck(url: string): string | undefined {
  let requested: string;
  try {
    requested = decodeURIComponent(url.split("?")[0]);
  } catch {
    // Malformed percent-encoding; not a deck we can serve.
    return undefined;
  }

  // Leading slashes would make an absolute path of their own; anything that
  // still resolves outside decks/ (traversal, a drive letter) fails the check
  // below rather than being pattern-matched here.
  const path = resolve(DECKS_DIR, requested.replace(/^[/\\]+/, ""));
  if (!path.startsWith(DECKS_DIR + sep)) return undefined;
  return existsSync(path) ? path : undefined;
}

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
        const path = resolveDeck(req.url ?? "");
        if (!path) return next();
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
 * Serves the test harness pages at / (single slide) and /thumbnails.html
 * (thumbnail list), plus the generated fixture decks (fixtures/*.pptx) as
 * static files at the server root (e.g. /basic.pptx).
 */
export default defineConfig({
  root: "harness",
  publicDir: "../fixtures",
  plugins: [serveLocalDecks()],
  // The thumbnail harness is the only JSX here and needs no refresh tooling,
  // so esbuild's automatic runtime replaces @vitejs/plugin-react.
  esbuild: { jsx: "automatic" },
  server: {
    fs: {
      // Paths are relative to the "harness" root. ".." covers this package
      // (incl. node_modules/.vite prebundled deps); "../../core" covers the
      // @diceui/pptx-core sources the workspace alias resolves to.
      allow: ["..", "../../core"],
    },
  },
});
