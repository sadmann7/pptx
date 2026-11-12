import { createReadStream, existsSync } from "node:fs";
import { resolve, sep } from "node:path";
import { defineConfig, type Plugin } from "vite";

const DECKS_DIR = resolve(import.meta.dirname, "decks");

// The containment check is what makes this safe: whatever the request asks for,
// the served file has to sit under decks/.
function resolveDeck(url: string): string | undefined {
  let requested: string;
  try {
    requested = decodeURIComponent(url.split("?")[0]);
  } catch {
    // Malformed percent-encoding can't name a servable deck.
    return undefined;
  }

  // Leading slashes would make an absolute path of their own; anything that
  // still resolves outside decks/ (traversal, a drive letter) fails the check
  // below rather than being pattern-matched here.
  const path = resolve(DECKS_DIR, requested.replace(/^[/\\]+/, ""));
  if (!path.startsWith(DECKS_DIR + sep)) return undefined;
  return existsSync(path) ? path : undefined;
}

// A real deck that reproduces a bug often cannot be committed; dropped in the
// gitignored decks/, it loads as ?file=decks/whatever.pptx.
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

// The bundle inlines the decode worker as a blob, while Vite spawns it from
// source, which is much slower to start. Requires `pnpm -F @diceui/pptx-core
// build`, or the request 404s.
function serveBuiltFontEntry(): Plugin {
  const file = resolve(import.meta.dirname, "..", "core", "dist", "fonts", "index.mjs");
  return {
    name: "e2e-serve-built-font-entry",
    apply: "serve",
    configureServer(server) {
      server.middlewares.use("/core-dist/fonts.mjs", (_req, res, next) => {
        if (!existsSync(file)) return next();
        res.setHeader("Content-Type", "text/javascript");
        createReadStream(file).pipe(res);
      });
    },
  };
}

export default defineConfig({
  root: "harness",
  publicDir: "../fixtures",
  plugins: [serveLocalDecks(), serveBuiltFontEntry()],
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
