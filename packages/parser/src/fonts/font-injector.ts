/**
 * Inject embedded PPTX fonts into the DOM.
 *
 * Decompression (LZCOMP + adaptive Huffman per font part) is CPU-heavy, so
 * unique font parts are decoded in parallel across a Web Worker pool. The
 * main thread only registers the resulting TrueType binaries with the
 * document via the FontFace API, which is cheap.
 *
 * Decks routinely embed dozens of font parts (every family x weight x style)
 * while the first-rendered slide uses only a few. Callers can pass
 * `priorityTypefaces` so `ready` resolves as soon as those are registered;
 * the remaining fonts keep decoding in the background and swap in when
 * registered (renderers re-run text autofit on `document.fonts` arrival).
 *
 * When Workers are unavailable (SSR, worker load failure), decoding falls
 * back to the main thread, yielding to the event loop between fonts.
 */

import type {
  EmbeddedFontEntry,
  EmbeddedFontVariant,
  PresentationData,
} from "../model/presentation";
import { decodeEmbeddedFont, toStandaloneArrayBuffer } from "./font-decode";
import type { FontWorkerRequest, FontWorkerResponse } from "./font-worker";

export interface FontInjectionHandle {
  /**
   * Resolves when the priority fonts are registered (or skipped). When no
   * `priorityTypefaces` were given, this waits for every embedded font.
   */
  ready: Promise<void>;
  /** Resolves when every embedded font has been registered (or skipped). */
  complete: Promise<void>;
  dispose(): void;
}

export interface InjectEmbeddedFontsOptions {
  /**
   * Typeface names that block `ready`. Anything else decodes in the
   * background after them. Names must match `EmbeddedFontEntry.typeface`.
   */
  priorityTypefaces?: ReadonlySet<string>;

  /**
   * Called after each font part finishes (decoded and registered, or
   * skipped on failure). `done` counts finished parts, `total` is the
   * number of unique font parts in the deck.
   */
  onProgress?: (done: number, total: number) => void;
}

const MAX_WORKERS = 6;

const VARIANTS: {
  key: keyof Pick<EmbeddedFontEntry, "regular" | "bold" | "italic" | "boldItalic">;
  weight: string;
  style: string;
}[] = [
  { key: "regular", weight: "normal", style: "normal" },
  { key: "bold", weight: "bold", style: "normal" },
  { key: "italic", weight: "normal", style: "italic" },
  { key: "boldItalic", weight: "bold", style: "italic" },
];

interface FontTask {
  typeface: string;
  weight: string;
  style: string;
  variant: EmbeddedFontVariant;
}

interface DecodeJob {
  path: string;
  bytes: Uint8Array;
  fontKey?: string;
}

/** Yield to the event loop so rendering and input stay responsive. */
function nextTick(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

// ── Worker pool ─────────────────────────────────────────────────────

/**
 * Decode jobs across a Web Worker pool, invoking `onDecoded` as each result
 * arrives (in queue order per worker, priority jobs first in the queue).
 * Jobs whose worker dies (e.g. the worker script failed to load) are left
 * un-decoded; the caller runs a main-thread fallback for anything missing.
 */
async function decodeWithWorkerPool(
  jobs: DecodeJob[],
  isDisposed: () => boolean,
  onDecoded: (path: string, buffer: ArrayBuffer | null) => void,
): Promise<Set<string>> {
  const decodedPaths = new Set<string>();
  const queue = [...jobs];

  const concurrency =
    typeof navigator !== "undefined" && navigator.hardwareConcurrency
      ? navigator.hardwareConcurrency
      : 4;
  const poolSize = Math.max(1, Math.min(jobs.length, concurrency - 1, MAX_WORKERS));

  const workers: Worker[] = [];
  for (let i = 0; i < poolSize; i++) {
    workers.push(new Worker(new URL("./font-worker.ts", import.meta.url), { type: "module" }));
  }

  await Promise.all(
    workers.map(
      (worker) =>
        new Promise<void>((resolve) => {
          const takeNext = (): void => {
            if (isDisposed()) {
              resolve();
              return;
            }
            const job = queue.shift();
            if (!job) {
              resolve();
              return;
            }
            const request: FontWorkerRequest = {
              path: job.path,
              // Copy: the original buffer stays usable by the rest of the app.
              bytes: toStandaloneArrayBuffer(job.bytes),
              fontKey: job.fontKey,
            };
            worker.postMessage(request, [request.bytes]);
          };

          worker.onmessage = (event: MessageEvent<FontWorkerResponse>) => {
            decodedPaths.add(event.data.path);
            onDecoded(event.data.path, event.data.buffer);
            takeNext();
          };
          // Fires when the worker script itself fails to load or crashes.
          // Stop using this worker; unfinished jobs fall back to the caller.
          worker.onerror = () => {
            resolve();
          };

          takeNext();
        }),
    ),
  );

  for (const worker of workers) {
    worker.terminate();
  }

  return decodedPaths;
}

// ── Public API ──────────────────────────────────────────────────────

export function injectEmbeddedFonts(
  presentation: PresentationData,
  options?: InjectEmbeddedFontsOptions,
): FontInjectionHandle {
  const noop: FontInjectionHandle = {
    ready: Promise.resolve(),
    complete: Promise.resolve(),
    dispose() {},
  };

  if (!presentation.embeddedFonts || presentation.embeddedFonts.length === 0) return noop;
  if (typeof document === "undefined" || typeof FontFace === "undefined") return noop;

  const tasks: FontTask[] = [];
  for (const entry of presentation.embeddedFonts) {
    for (const { key, weight, style } of VARIANTS) {
      const variant = entry[key];
      if (variant) tasks.push({ typeface: entry.typeface, weight, style, variant });
    }
  }
  if (tasks.length === 0) return noop;

  // Same .fntdata part can back multiple typeface entries — decode once.
  const jobByPath = new Map<string, DecodeJob>();
  const tasksByPath = new Map<string, FontTask[]>();
  for (const task of tasks) {
    const path = task.variant.path;
    const forPath = tasksByPath.get(path);
    if (forPath) {
      forPath.push(task);
    } else {
      tasksByPath.set(path, [task]);
    }
    if (jobByPath.has(path)) continue;
    const bytes = presentation.fonts.get(path);
    if (!bytes || bytes.length === 0) continue;
    jobByPath.set(path, { path, bytes, fontKey: task.variant.fontKey });
  }
  if (jobByPath.size === 0) return noop;

  // A part is priority when any of its typefaces is priority. Without an
  // explicit priority set, everything is priority (ready === complete).
  const priority = options?.priorityTypefaces;
  const isPriorityPath = (path: string): boolean =>
    !priority || (tasksByPath.get(path) ?? []).some((task) => priority.has(task.typeface));

  // Priority parts decode first.
  const jobs = [...jobByPath.values()].sort(
    (a, b) => Number(isPriorityPath(b.path)) - Number(isPriorityPath(a.path)),
  );
  let pendingPriority = jobs.filter((job) => isPriorityPath(job.path)).length;

  const registered: FontFace[] = [];
  let disposed = false;
  const isDisposed = () => disposed;

  let resolveReady!: () => void;
  const ready = new Promise<void>((resolve) => {
    resolveReady = resolve;
  });
  if (pendingPriority === 0) resolveReady();

  const totalParts = jobs.length;
  let partsDone = 0;

  /** Register every typeface variant backed by a decoded part. */
  async function registerPath(path: string, buffer: ArrayBuffer | null): Promise<void> {
    if (buffer) {
      for (const task of tasksByPath.get(path) ?? []) {
        if (disposed) return;
        try {
          const face = new FontFace(task.typeface, buffer, {
            weight: task.weight,
            style: task.style,
          });
          await face.load();
          if (disposed) return;
          document.fonts.add(face);
          registered.push(face);
        } catch {
          // Invalid font data — skip this variant, text falls back.
        }
      }
    }
    partsDone += 1;
    options?.onProgress?.(partsDone, totalParts);
    if (isPriorityPath(path) && --pendingPriority === 0) {
      resolveReady();
    }
  }

  const registrations: Promise<void>[] = [];
  const onDecoded = (path: string, buffer: ArrayBuffer | null): void => {
    registrations.push(registerPath(path, buffer));
  };

  const complete = (async () => {
    let decodedPaths = new Set<string>();

    if (typeof Worker !== "undefined") {
      try {
        decodedPaths = await decodeWithWorkerPool(jobs, isDisposed, onDecoded);
      } catch {
        decodedPaths = new Set();
      }
    }

    // Main-thread fallback for anything the pool did not decode
    // (Workers unavailable, worker script failed to load, or died mid-run).
    for (const job of jobs) {
      if (disposed) return;
      if (decodedPaths.has(job.path)) continue;
      const decoded = decodeEmbeddedFont(job.bytes, job.fontKey);
      onDecoded(job.path, decoded ? toStandaloneArrayBuffer(decoded) : null);
      await nextTick();
    }

    await Promise.all(registrations);
    // Safety net: never leave `ready` pending (e.g. after dispose).
    resolveReady();
  })();

  return {
    ready,
    complete,
    dispose() {
      disposed = true;
      resolveReady();
      for (const face of registered) {
        document.fonts.delete(face);
      }
      registered.length = 0;
    },
  };
}

/**
 * Match raw OOXML part sources (a slide plus its layout/master, which text
 * inherits typefaces from) against the deck's embedded typefaces, returning
 * the set actually referenced. Theme major/minor fonts are always included,
 * since text reaches them via `+mj-lt`/`+mn-lt` references.
 *
 * Used to prioritize first-slide fonts so `ready` does not wait for every
 * embedded family in the deck. Late fonts still swap in when registered.
 */
export function collectPriorityTypefaces(
  presentation: PresentationData,
  sources: ReadonlyArray<string | undefined>,
): Set<string> | undefined {
  const embedded = presentation.embeddedFonts;
  const xmlSources = sources.filter((s): s is string => !!s);
  if (!embedded || embedded.length === 0 || xmlSources.length === 0) return undefined;

  const priority = new Set<string>();
  const themeFonts = new Set<string>();
  for (const theme of presentation.themes.values()) {
    if (theme.majorFont?.latin) themeFonts.add(theme.majorFont.latin);
    if (theme.minorFont?.latin) themeFonts.add(theme.minorFont.latin);
  }

  for (const entry of embedded) {
    if (
      themeFonts.has(entry.typeface) ||
      xmlSources.some((xml) => xml.includes(`typeface="${entry.typeface}"`))
    ) {
      priority.add(entry.typeface);
    }
  }
  return priority;
}
