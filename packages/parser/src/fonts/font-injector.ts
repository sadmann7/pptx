/**
 * Inject embedded PPTX fonts into the DOM.
 *
 * Decompression (LZCOMP + adaptive Huffman per font part) is CPU-heavy, so
 * unique font parts are decoded in parallel across a Web Worker pool. The
 * main thread only registers the resulting TrueType binaries with the
 * document via the FontFace API, which is cheap.
 *
 * When Workers are unavailable (SSR, worker load failure), decoding falls
 * back to the main thread, yielding to the event loop between fonts.
 *
 * Callers should await `handle.ready` before rendering slides so text is
 * measured and painted with the embedded fonts (no fallback-font layout
 * shift).
 */

import type {
  EmbeddedFontEntry,
  EmbeddedFontVariant,
  PresentationData,
} from "../model/presentation";
import { decodeEmbeddedFont, toStandaloneArrayBuffer } from "./font-decode";
import type { FontWorkerRequest, FontWorkerResponse } from "./font-worker";

export interface FontInjectionHandle {
  /** Resolves when every embedded font has been registered (or skipped). */
  ready: Promise<void>;
  dispose(): void;
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
 * Decode jobs across a Web Worker pool. Jobs whose worker dies (e.g. the
 * worker script failed to load) are left un-decoded; the caller runs a
 * main-thread fallback for anything missing from the result map.
 */
async function decodeWithWorkerPool(
  jobs: DecodeJob[],
  isDisposed: () => boolean,
): Promise<Map<string, ArrayBuffer | null>> {
  const results = new Map<string, ArrayBuffer | null>();
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
            results.set(event.data.path, event.data.buffer);
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

  return results;
}

// ── Public API ──────────────────────────────────────────────────────

export function injectEmbeddedFonts(presentation: PresentationData): FontInjectionHandle {
  const noop: FontInjectionHandle = { ready: Promise.resolve(), dispose() {} };

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
  for (const task of tasks) {
    if (jobByPath.has(task.variant.path)) continue;
    const bytes = presentation.fonts.get(task.variant.path);
    if (!bytes || bytes.length === 0) continue;
    jobByPath.set(task.variant.path, {
      path: task.variant.path,
      bytes,
      fontKey: task.variant.fontKey,
    });
  }
  if (jobByPath.size === 0) return noop;

  const registered: FontFace[] = [];
  let disposed = false;
  const isDisposed = () => disposed;

  const ready = (async () => {
    let buffers: Map<string, ArrayBuffer | null>;

    if (typeof Worker !== "undefined") {
      try {
        buffers = await decodeWithWorkerPool([...jobByPath.values()], isDisposed);
      } catch {
        buffers = new Map();
      }
    } else {
      buffers = new Map();
    }

    // Main-thread fallback for anything the pool did not decode
    // (Workers unavailable, worker script failed to load, or died mid-run).
    for (const job of jobByPath.values()) {
      if (disposed) return;
      if (buffers.has(job.path)) continue;
      const decoded = decodeEmbeddedFont(job.bytes, job.fontKey);
      buffers.set(job.path, decoded ? toStandaloneArrayBuffer(decoded) : null);
      await nextTick();
    }

    // Register a FontFace per typeface variant. Cheap relative to decode.
    for (const task of tasks) {
      if (disposed) return;
      const buffer = buffers.get(task.variant.path);
      if (!buffer) continue;
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
  })();

  return {
    ready,
    dispose() {
      disposed = true;
      for (const face of registered) {
        document.fonts.delete(face);
      }
      registered.length = 0;
    },
  };
}
