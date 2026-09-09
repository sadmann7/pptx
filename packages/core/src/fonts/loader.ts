/**
 * Loads embedded PPTX fonts and registers them with the document.
 *
 * Only MTX-compressed parts (LZCOMP + adaptive Huffman, a few ms each) are
 * worth moving. A deck carrying one sends its parts to a Web Worker pool, which
 * decodes unique parts in parallel and leaves the main thread to register the
 * resulting TrueType binaries via the FontFace API, which is cheap. A deck
 * whose parts are all uncompressed decodes in microseconds, so a pool would
 * cost more to start than it could save, and it is decoded here instead.
 *
 * A deck can embed a part per family x weight x style while the first-rendered
 * slide uses only a few. Callers can pass `priorityTypefaces` so `ready`
 * resolves as soon as those are registered; the remaining fonts keep decoding
 * in the background and swap in when registered (renderers re-run text autofit
 * on `document.fonts` arrival).
 *
 * Main-thread decoding gives the event loop a turn whenever it has held it long
 * enough to risk a frame, rather than after every part: a turn of the loop
 * costs about as much as decoding an uncompressed part, so yielding per part
 * would be most of what decoding such a deck spends.
 */

import type {
  EmbeddedFontEntry,
  EmbeddedFontVariant,
  PresentationData,
} from "../model/presentation";
import { copyToArrayBuffer, decodeEmbeddedFontPart, getIsCompressedFont } from "./decode";
import type { MtxErrorCode } from "./mtx";
import type { FontWorkerRequest, FontWorkerResponse } from "./worker";
import { createFontWorker } from "./worker-host";

const MAX_WORKER_COUNT = 6;

// Max time main-thread decoding holds the event loop before yielding (~1 frame at 60Hz).
const SLICE_MS = 5;

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

export interface EmbeddedFontsHandle {
  /**
   * Resolves when the priority fonts are registered (or skipped). When no
   * `priorityTypefaces` were given, this waits for every embedded font.
   */
  ready: Promise<void>;
  /** Resolves when every embedded font has been registered (or skipped). */
  complete: Promise<void>;
  dispose(): void;
}

/** An embedded font that could not be used, reported through `onError`. */
export interface EmbeddedFontError {
  /** Package path of the `.fntdata` part that failed. */
  path: string;
  /** The typefaces left without this part. */
  typefaces: string[];
  /**
   * `"missing"`: the deck references a part it does not contain.
   * `"decode"`: the part is not a font this decoder can read.
   * `"register"`: the decoded bytes were rejected by the FontFace API.
   */
  stage: "missing" | "decode" | "register";
  message: string;
  /** Set when the MTX decoder identified the failure. */
  code?: MtxErrorCode;
}

export interface LoadEmbeddedFontsOptions {
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
  /**
   * Called for each embedded font that could not be used. Loading continues
   * either way and the affected text renders with a fallback typeface, so this
   * exists to make an otherwise silent gap observable.
   *
   * `"missing"` reports trigger while `loadEmbeddedFonts` is still building its
   * job list, before it returns.
   */
  onError?: (error: EmbeddedFontError) => void;
}

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

interface DecodeFailure {
  message: string;
  code?: MtxErrorCode;
}

type OnDecoded = (path: string, buffer: ArrayBuffer | null, failure?: DecodeFailure) => void;

/**
 * Decode jobs across a Web Worker pool, invoking `onDecoded` as each result
 * arrives (in queue order per worker, priority jobs first in the queue).
 * Jobs whose worker dies (e.g. the worker script failed to load) are left
 * un-decoded; the caller runs a main-thread fallback for anything missing.
 */
async function decodeWithWorkerPool(
  jobs: DecodeJob[],
  isDisposed: () => boolean,
  onDecoded: OnDecoded,
): Promise<Set<string>> {
  const decodedPaths = new Set<string>();
  const queue = [...jobs];

  const concurrency =
    typeof navigator !== "undefined" && navigator.hardwareConcurrency
      ? navigator.hardwareConcurrency
      : 4;
  const poolSize = Math.max(1, Math.min(jobs.length, concurrency - 1, MAX_WORKER_COUNT));

  const workers: Worker[] = [];
  for (let i = 0; i < poolSize; i++) {
    workers.push(createFontWorker());
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
              bytes: copyToArrayBuffer(job.bytes),
              fontKey: job.fontKey,
            };
            worker.postMessage(request, [request.bytes]);
          };

          worker.onmessage = (event: MessageEvent<FontWorkerResponse>) => {
            const { path, buffer, message, code } = event.data;
            decodedPaths.add(path);
            onDecoded(
              path,
              buffer,
              buffer ? undefined : { message: message ?? "Decode failed", code },
            );
            takeNext();
          };
          // Triggers when the worker script itself fails to load or crashes.
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

export function loadEmbeddedFonts(
  presentation: PresentationData,
  options?: LoadEmbeddedFontsOptions,
): EmbeddedFontsHandle {
  const noop: EmbeddedFontsHandle = {
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

  // Same .fntdata part can back multiple typeface entries; decode once.
  const jobByPath = new Map<string, DecodeJob>();
  const tasksByPath = new Map<string, FontTask[]>();
  const missingPaths = new Set<string>();
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
    if (!bytes) {
      missingPaths.add(path);
      continue;
    }
    // A part that is present but empty stays a job, so the decoder reports it
    // as the unreadable payload it is rather than as one the deck never had.
    jobByPath.set(path, { path, bytes, fontKey: task.variant.fontKey });
  }

  const typefacesForPath = (path: string): string[] => [
    ...new Set((tasksByPath.get(path) ?? []).map((task) => task.typeface)),
  ];

  for (const path of missingPaths) {
    options?.onError?.({
      path,
      typefaces: typefacesForPath(path),
      stage: "missing",
      message: "Font part is not present in the package",
    });
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

  const reportError = (error: EmbeddedFontError): void => {
    if (disposed) return;
    options?.onError?.(error);
  };

  /** Register every typeface variant backed by a decoded part. */
  async function registerPath(
    path: string,
    buffer: ArrayBuffer | null,
    failure?: DecodeFailure,
  ): Promise<void> {
    if (buffer) {
      // Every variant of a part shares its bytes, so a rejection is the part's
      // and not one variant's. They are collected into a single report rather
      // than one per family, to match how the other stages read.
      const rejected: string[] = [];
      let rejection: string | undefined;
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
        } catch (error) {
          // Invalid font data: skip this variant, text falls back.
          rejected.push(task.typeface);
          rejection ??= error instanceof Error ? error.message : String(error);
        }
      }
      if (rejected.length > 0) {
        reportError({
          path,
          typefaces: [...new Set(rejected)],
          stage: "register",
          message: rejection ?? "FontFace rejected the decoded font",
        });
      }
    } else {
      reportError({
        path,
        typefaces: typefacesForPath(path),
        stage: "decode",
        message: failure?.message ?? "Decode failed",
        code: failure?.code,
      });
    }
    partsDone += 1;
    options?.onProgress?.(partsDone, totalParts);
    if (isPriorityPath(path) && --pendingPriority === 0) {
      resolveReady();
    }
  }

  const registrations: Promise<void>[] = [];
  const onDecoded: OnDecoded = (path, buffer, failure) => {
    registrations.push(registerPath(path, buffer, failure));
  };

  const complete = (async () => {
    let decodedPaths = new Set<string>();

    // Starting a pool costs more than decoding uncompressed parts does, so one
    // compressed part sends the whole deck to the pool and none skips it.
    const hasCompressedPart = jobs.some((job) => getIsCompressedFont(job.bytes, job.fontKey));

    if (hasCompressedPart && typeof Worker !== "undefined") {
      try {
        decodedPaths = await decodeWithWorkerPool(jobs, isDisposed, onDecoded);
      } catch {
        decodedPaths = new Set();
      }
    }

    // Whatever the pool did not decode: uncompressed payloads, or parts left
    // behind when it was skipped (SSR) or failed mid-run.
    let sliceStart = performance.now();
    for (const job of jobs) {
      if (disposed) return;
      if (decodedPaths.has(job.path)) continue;
      const result = decodeEmbeddedFontPart(job.bytes, job.fontKey);
      if (result.ok) {
        onDecoded(job.path, copyToArrayBuffer(result.bytes));
      } else {
        onDecoded(job.path, null, { message: result.message, code: result.code });
      }
      if (performance.now() - sliceStart < SLICE_MS) continue;
      await new Promise((resolve) => setTimeout(resolve, 0));
      sliceStart = performance.now();
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
 * Pick out the embedded typefaces that the given slides actually reference,
 * for use as `priorityTypefaces`.
 *
 * `slideXml` holds the raw OOXML of the slides about to be shown, each
 * accompanied by its layout and master, since text inherits typefaces from
 * those. Theme major/minor fonts are always included, because text reaches
 * them indirectly via `+mj-lt`/`+mn-lt` references rather than by name.
 *
 * Prioritizing these keeps `ready` from waiting on every embedded family in
 * the deck. The rest still swap in once they are registered.
 */
export function findPriorityTypefaces(
  presentation: PresentationData,
  slideXml: ReadonlyArray<string | undefined>,
): Set<string> | undefined {
  const embedded = presentation.embeddedFonts;
  const xmlSources = slideXml.filter((xml): xml is string => !!xml);
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
