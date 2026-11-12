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

// Max time main-thread decoding holds the event loop before yielding (~1 frame at
// 60Hz). Yielding per part would cost about as much as decoding an uncompressed one.
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
   * Resolves when the priority fonts are registered or skipped, or every embedded font
   * when no `priorityTypefaces` were given.
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
   * Typeface names that block `ready`, matching `EmbeddedFontEntry.typeface`; the rest
   * decode in the background after them.
   */
  priorityTypefaces?: ReadonlySet<string>;
  /**
   * Called after each font part is registered or skipped, with `total` counting the deck's
   * unique font parts.
   */
  onProgress?: (done: number, total: number) => void;
  /**
   * Called for each embedded font that falls back to another typeface; `"missing"` reports
   * fire before `loadEmbeddedFonts` returns.
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

// Results arrive in queue order per worker. Jobs whose worker dies stay un-decoded,
// and the caller runs a main-thread fallback for them.
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
          // Fires when the worker script fails to load or crashes; unfinished
          // jobs fall back to the caller.
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

  const priority = options?.priorityTypefaces;
  const isPriorityPath = (path: string): boolean =>
    !priority || (tasksByPath.get(path) ?? []).some((task) => priority.has(task.typeface));

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

  async function registerPath(
    path: string,
    buffer: ArrayBuffer | null,
    failure?: DecodeFailure,
  ): Promise<void> {
    if (buffer) {
      // All variants share part bytes, so a rejection is reported once per part.
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
    // `ready` must never stay pending, including after dispose.
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
 * Picks the embedded typefaces the given slides reference, for `priorityTypefaces`. Pass
 * each slide's layout and master XML too, since text inherits typefaces from them.
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
