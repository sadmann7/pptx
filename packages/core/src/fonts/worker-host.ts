/**
 * Creates the embedded-font decode worker.
 *
 * In dev, this loads `./worker.ts` directly. For published builds (see `tsdown.config.ts`),
 * the worker source is inlined and spawned from a blob.
 */
export function createFontWorker(): Worker {
  return new Worker(new URL("./worker.ts", import.meta.url), { type: "module" });
}
