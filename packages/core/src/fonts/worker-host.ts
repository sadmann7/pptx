/**
 * Creates the embedded-font decode worker.
 *
 * This module is replaced at build time (see `tsdown.config.ts`): the published
 * build inlines `worker.ts` and its decoder tree as a bundled source string and
 * spawns the worker from a blob URL. The version below is what runs when the
 * package is consumed as source, where `./worker.ts` is a real resolvable file.
 *
 * A relative URL cannot survive publishing. Bundlers rewrite the paths around
 * it, rolldown has no worker handling to emit a matching chunk, and the
 * specifier would be left pointing at a source file that is not shipped.
 */
export function createFontWorker(): Worker {
  return new Worker(new URL("./worker.ts", import.meta.url), { type: "module" });
}
