export interface PdfjsOptions {
  /**
   * URL for the pdfjs ESM module, for example
   * `new URL('pdfjs-dist/build/pdf.min.mjs', import.meta.url).toString()`.
   */
  moduleUrl?: string;
  /**
   * URL for the pdfjs worker ESM module, for example
   * `new URL('pdfjs-dist/build/pdf.worker.min.mjs', import.meta.url).toString()`.
   */
  workerUrl?: string;
}

export type PdfjsConfig = PdfjsOptions | false;

const PDFJS_MODULE_SPECIFIER = "pdfjs-dist/build/pdf.min.mjs";
const PDFJS_WORKER_SPECIFIER = "pdfjs-dist/build/pdf.worker.min.mjs";

let _pdfjsUrl: string | null = null;
let _pdfWorkerUrl: string | null = null;

function resolveModuleUrl(specifier: string): string | null {
  try {
    const resolver = (import.meta as ImportMeta & { resolve?: (id: string) => string }).resolve;
    if (typeof resolver === "function") {
      return resolver(specifier);
    }
  } catch {
    // The host runtime does not expose import.meta.resolve, or cannot resolve pdfjs-dist.
  }
  return null;
}

function explicitUrl(config: PdfjsConfig | undefined, key: keyof PdfjsOptions): string | null {
  if (!config || typeof config !== "object") return null;
  const url = config[key];
  if (typeof url !== "string") return null;
  const trimmed = url.trim();
  return trimmed || null;
}

function getPdfjsUrl(): string | null {
  if (_pdfjsUrl !== null) return _pdfjsUrl;
  _pdfjsUrl = resolveModuleUrl(PDFJS_MODULE_SPECIFIER) ?? "";
  return _pdfjsUrl || null;
}

function getPdfWorkerUrl(): string | null {
  if (_pdfWorkerUrl !== null) return _pdfWorkerUrl;
  _pdfWorkerUrl = resolveModuleUrl(PDFJS_WORKER_SPECIFIER) ?? "";
  return _pdfWorkerUrl || null;
}

function resolvePdfjsUrl(config?: PdfjsConfig): string | null {
  if (config === false) return null;
  return explicitUrl(config, "moduleUrl") ?? getPdfjsUrl();
}

function resolvePdfWorkerUrl(config?: PdfjsConfig): string | null {
  if (config === false) return null;
  return explicitUrl(config, "workerUrl") ?? getPdfWorkerUrl();
}

/**
 * pdfjs keeps process-level state (GlobalWorkerOptions.workerSrc, PagesMapper)
 * that would clobber the host app's own pdfjs setup, so PDFs render only in
 * this worker, with its own pdfjs instance. There is no main-thread fallback.
 *
 * Receives `{ id, pdfData, width, height, pdfjsUrl, pdfWorkerUrl }` and posts
 * back `{ id, blob }` or `{ id, error }`.
 */
const WORKER_SRC = /* js */ `
let pdfjsLib = null;

self.onmessage = async (e) => {
  const { id, pdfData, width, height, pdfjsUrl, pdfWorkerUrl } = e.data;
  try {
    if (!pdfjsLib) {
      pdfjsLib = await import(pdfjsUrl);
      pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
    }

    const doc = await pdfjsLib.getDocument({ data: pdfData }).promise;
    try {
      if (doc.numPages < 1) {
        self.postMessage({ id, error: 'no pages' });
        return;
      }
      const page = await doc.getPage(1);
      const vp = page.getViewport({ scale: 1 });
      const scale = Math.max(width / vp.width, height / vp.height);
      const svp = page.getViewport({ scale });

      const canvas = new OffscreenCanvas(Math.ceil(svp.width), Math.ceil(svp.height));
      const ctx = canvas.getContext('2d', { alpha: true });
      await page.render({ canvasContext: ctx, viewport: svp, background: 'rgba(0,0,0,0)' }).promise;

      const blob = await canvas.convertToBlob({ type: 'image/png' });
      self.postMessage({ id, blob });
    } finally {
      doc.destroy();
    }
  } catch (error) {
    self.postMessage({ id, error: String(error) });
  }
};
`;

interface WorkerState {
  worker: Worker | null;
  failed: boolean;
  msgId: number;
  pending: Map<number, { resolve: (b: Blob | null) => void; reject: (e: Error) => void }>;
}

const _workers = new Map<string, WorkerState>();

function getWorkerState(workerKey: string): WorkerState {
  let state = _workers.get(workerKey);
  if (!state) {
    state = {
      worker: null,
      failed: false,
      msgId: 0,
      pending: new Map(),
    };
    _workers.set(workerKey, state);
  }
  return state;
}

function getWorker(workerKey: string): Worker | null {
  const state = getWorkerState(workerKey);
  if (state.failed) return null;
  if (state.worker) return state.worker;

  try {
    const blob = new Blob([WORKER_SRC], { type: "text/javascript" });
    const url = URL.createObjectURL(blob);
    try {
      state.worker = new Worker(url, { type: "module" });
    } finally {
      URL.revokeObjectURL(url);
    }

    state.worker.onmessage = (e: MessageEvent) => {
      const { id, blob, error } = e.data;
      const entry = state.pending.get(id);
      if (!entry) return;
      state.pending.delete(id);
      if (error) {
        entry.resolve(null); // A worker-side error counts as "no result"
      } else {
        entry.resolve(blob ?? null);
      }
    };

    state.worker.onerror = () => {
      // Worker failed to initialize (e.g. module import blocked by CSP).
      state.failed = true;
      state.worker = null;
      for (const [, entry] of state.pending) {
        entry.resolve(null);
      }
      state.pending.clear();
    };

    return state.worker;
  } catch {
    state.failed = true;
    return null;
  }
}

function renderInWorker(
  pdfData: Uint8Array,
  width: number,
  height: number,
  pdfjsUrl: string,
  pdfWorkerUrl: string,
): Promise<Blob | null> {
  return new Promise((resolve) => {
    const workerKey = `${pdfjsUrl}\n${pdfWorkerUrl}`;
    const state = getWorkerState(workerKey);
    const worker = getWorker(workerKey);
    if (!worker) {
      resolve(null);
      return;
    }

    const id = ++state.msgId;
    state.pending.set(id, {
      resolve,
      reject: () => resolve(null),
    });

    // Transfer a copy so the caller's buffer stays attached.
    const copy = pdfData.slice();
    worker.postMessage({ id, pdfData: copy, width, height, pdfjsUrl, pdfWorkerUrl }, [copy.buffer]);

    setTimeout(() => {
      if (state.pending.has(id)) {
        state.pending.delete(id);
        resolve(null);
      }
    }, 15000);
  });
}

/**
 * Renders page 1 of a PDF to a blob URL in a worker, or returns null when
 * rendering fails or Worker/OffscreenCanvas are unavailable.
 */
export async function renderPdfToImage(
  pdfData: Uint8Array,
  width: number,
  height: number,
  pdfjs?: PdfjsConfig,
): Promise<string | null> {
  const pdfjsUrl = resolvePdfjsUrl(pdfjs);
  const pdfWorkerUrl = resolvePdfWorkerUrl(pdfjs);

  if (
    !pdfjsUrl ||
    !pdfWorkerUrl ||
    typeof OffscreenCanvas === "undefined" ||
    typeof Worker === "undefined"
  ) {
    return null;
  }

  try {
    const blob = await renderInWorker(pdfData, width, height, pdfjsUrl, pdfWorkerUrl);
    if (blob) return URL.createObjectURL(blob);
  } catch {
    // A failed render leaves the picture blank, not the slide broken.
  }

  return null;
}
