// Dev loads `./worker.ts` directly; published builds inline the worker source (see
// `tsdown.config.ts`) and spawn it from a blob.
export function createFontWorker(): Worker {
  return new Worker(new URL("./worker.ts", import.meta.url), { type: "module" });
}
