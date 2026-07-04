import type { FontInjectionHandle, PresentationData, SlideData } from "@diceui/pptx-parser";
import { buildPresentation, materializeSlideNodes, parseZipLazyMedia } from "@diceui/pptx-parser";

const INITIAL_STATE: PresentationState = {
  status: "idle",
  presentation: null,
  activeSlideId: null,
  zoom: 1,
  progress: 0,
  error: null,
};

const ABORT_ERROR = new DOMException("Superseded by a newer load", "AbortError");

/** Accepted input formats for `store.load()` and the `file` prop. */
export type PreviewInput = ArrayBuffer | Uint8Array | Blob | File;

/**
 * Lifecycle status of the presentation viewer.
 *
 * - `"idle"`: no file loaded yet (initial state, or after `reset()`).
 * - `"loading"`: parsing is in progress.
 * - `"ready"`: a presentation is loaded and slides are available.
 * - `"error"`: parsing failed; inspect `PresentationState.error`.
 */
export type PresentationStatus = "idle" | "loading" | "ready" | "error";

/** Full snapshot of the presentation viewer state. */
export interface PresentationState {
  /** Current lifecycle status. */
  status: PresentationStatus;

  /** Parsed presentation data, or `null` when none is loaded. */
  presentation: PresentationData | null;

  /**
   * Stable identity of the active slide (`SlideData.id`,
   * e.g. `"ppt/slides/slide3.xml"`). `null` when no presentation is loaded.
   *
   * Using a stable id instead of a positional index means reordering,
   * inserting, or deleting slides never silently redirects the viewer
   * to the wrong slide.
   */
  activeSlideId: string | null;

  /**
   * Current zoom level, where `1` equals 100%.
   *
   * @default 1
   */
  zoom: number;

  /**
   * Parse progress from `0` to `100`.
   * Only meaningful while `status === "loading"`.
   */
  progress: number;

  /** Error produced by the last failed `load()` call, or `null`. */
  error: Error | null;
}

/** Store returned by `useCreatePresentationStore` for controlled usage. */
export interface PresentationStore {
  /** Returns the current state snapshot. Non-reactive: use `subscribe` to watch for changes. */
  getState: () => PresentationState;
  /**
   * Subscribe to state changes.
   *
   * ```ts
   * const unsubscribe = store.subscribe(() => {
   *   console.log(store.getState().status);
   * });
   * // Later:
   * unsubscribe();
   * ```
   *
   * @returns A function that removes the subscription when called.
   */
  subscribe: (listener: () => void) => () => void;

  /**
   * Parse and display a presentation file.
   *
   * @returns The parsed `PresentationData` on success.
   * Rejects with a `DOMException` named `"AbortError"` if a newer `load()`
   * call superseded this one before it completed: callers can safely ignore
   * that error.
   *
   * ```ts
   * try {
   *   const presentation = await store.load(file);
   *   console.log("slides:", presentation.slides.length);
   * } catch (err) {
   *   if (err instanceof DOMException && err.name === "AbortError") return;
   *   console.error(err);
   * }
   * ```
   */
  load: (
    input: PreviewInput,
    options?: {
      /**
       * 0-based index of the slide to navigate to after a successful parse.
       * Also accepts a resolver called with the parsed slides, useful when the
       * target index depends on content.
       *
       * @default 0
       *
       * ```ts
       * // Static
       * store.load(file, { defaultSlideIndex: 2 });
       *
       * // Dynamic: last slide
       * store.load(file, { defaultSlideIndex: (slides) => slides.length - 1 });
       *
       * // Dynamic: by id
       * store.load(file, {
       *   defaultSlideIndex: (slides) => slides.findIndex(s => s.id === savedId),
       * });
       * ```
       */
      defaultSlideIndex?: number | ((slides: SlideData[]) => number);

      /**
       * When `true`, only the active slide's contents are parsed during
       * `load()`; other slides are parsed when first rendered or navigated
       * to. Makes time-to-first-slide near-constant regardless of deck size.
       * Set to `false` to parse every slide during `load()`.
       *
       * The store always materializes the active slide before notifying
       * subscribers, so `SlideData.nodes` of the current slide is reliable.
       * Non-active slides expose empty `nodes` until they are materialized.
       *
       * @default true
       */
      lazy?: boolean;

      /**
       * When `true`, fonts embedded in the PPTX are decoded and registered
       * with the browser before the store transitions to `"ready"`. This
       * eliminates font-swap layout shifts (FOUT) at the cost of a brief
       * extra loading phase while workers decode the font binaries.
       * Set to `false` to skip embedded font loading entirely (faster load,
       * but text may render with fallback fonts).
       *
       * @default true
       */
      embedFonts?: boolean;
    },
  ) => Promise<PresentationData>;

  /**
   * Reset the viewer to its initial idle state and cancel any in-flight load.
   *
   * ```ts
   * store.reset();
   * ```
   */
  reset: () => void;

  /**
   * Navigate to a slide by its stable id (`SlideData.id`).
   * No-op if the id is not found in the current presentation.
   */
  goTo: (slideId: string) => void;

  /**
   * Navigate to a slide by its 0-based index.
   * Clamps to the valid range; no-op when no presentation is loaded.
   */
  goToIndex: (index: number) => void;

  /** Navigate to the next slide. No-op when already on the last slide. */
  next: () => void;

  /** Navigate to the previous slide. No-op when already on the first slide. */
  prev: () => void;

  /**
   * Set the zoom level directly.
   * Clamped to `[0.1, 4]`.
   *
   * @default 1
   */
  setZoom: (zoom: number) => void;

  /**
   * Increase zoom by `step`.
   *
   * @default step 0.25
   */
  zoomIn: (step?: number) => void;

  /**
   * Decrease zoom by `step`.
   *
   * @default step 0.25
   */
  zoomOut: (step?: number) => void;

  /**
   * Fit the presentation to a container by computing the largest zoom level
   * that keeps all slides fully visible.
   *
   * ```ts
   * store.fitTo(containerWidth, containerHeight, 32);
   * ```
   */
  fitTo: (containerWidth: number, containerHeight: number, padding?: number) => void;

  /**
   * Returns the 0-based index of a slide by its stable id, or `-1` if not found.
   *
   * O(1) lookup via an internal id→index map. Safe to call inside
   * `useSyncExternalStore` selectors without the O(N) cost of `findIndex`.
   */
  getSlideIndex: (slideId: string) => number;

  /**
   * Returns the 0-based index of the active slide, or `-1` when none is active.
   */
  getActiveSlideIndex: () => number;

  /**
   * Returns the active `SlideData`, or `null` when none is active.
   */
  getActiveSlide: () => PresentationData["slides"][number] | null;

  /** Returns `true` when there is a next slide to navigate to. */
  canGoNext: () => boolean;

  /** Returns `true` when there is a previous slide to navigate to. */
  canGoPrev: () => boolean;
}

async function normalizeInput(input: PreviewInput): Promise<ArrayBuffer> {
  if (input instanceof ArrayBuffer) return input;

  if (input instanceof Uint8Array) {
    const buffer = input.buffer;
    if (
      buffer instanceof ArrayBuffer &&
      input.byteOffset === 0 &&
      input.byteLength === buffer.byteLength
    ) {
      return buffer;
    }
    return input.slice().buffer as ArrayBuffer;
  }

  if (input instanceof Blob) return input.arrayBuffer();

  throw new Error("Unsupported input type");
}

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}

export function createPresentationStore(): PresentationStore {
  let state: PresentationState = { ...INITIAL_STATE };
  const listeners = new Set<() => void>();
  let loadGeneration = 0;
  let fontInjection: FontInjectionHandle | undefined;

  let slideIndexById = new Map<string, number>();

  function rebuildSlideIndex(presentation: PresentationData | null): void {
    slideIndexById = new Map();
    if (!presentation) return;
    presentation.slides.forEach((slide, index) => {
      slideIndexById.set(slide.id, index);
    });
  }

  function emit(): void {
    for (const listener of listeners) listener();
  }

  function setState(
    next: Partial<PresentationState> | ((current: PresentationState) => Partial<PresentationState>),
  ): void {
    const patch = typeof next === "function" ? next(state) : next;
    state = { ...state, ...patch };
    if ("presentation" in patch) rebuildSlideIndex(state.presentation);
    emit();
  }

  function replaceState(next: PresentationState): void {
    state = next;
    rebuildSlideIndex(next.presentation);
    emit();
  }

  function getState(): PresentationState {
    return state;
  }

  function subscribe(listener: () => void): () => void {
    listeners.add(listener);
    return () => listeners.delete(listener);
  }

  async function load(
    input: PreviewInput,
    options?: {
      defaultSlideIndex?: number | ((slides: SlideData[]) => number);
      lazy?: boolean;
      embedFonts?: boolean;
    },
  ): Promise<PresentationData> {
    loadGeneration += 1;
    const gen = loadGeneration;

    fontInjection?.dispose();
    fontInjection = undefined;
    replaceState({ ...INITIAL_STATE, status: "loading", progress: 0 });

    try {
      const buffer = await normalizeInput(input);
      if (gen !== loadGeneration) throw ABORT_ERROR;
      setState({ progress: 30 });

      const files = await parseZipLazyMedia(buffer);
      if (gen !== loadGeneration) throw ABORT_ERROR;
      setState({ progress: 70 });

      // Lazy by default: defer per-slide node parsing until a slide is
      // rendered or navigated to, so load time stays flat for large decks.
      const presentation = buildPresentation(files, { lazy: options?.lazy ?? true });
      if (gen !== loadGeneration) throw ABORT_ERROR;
      setState({ progress: 85 });

      const defaultSlideIndex = options?.defaultSlideIndex;
      const requestedIndex =
        typeof defaultSlideIndex === "function"
          ? defaultSlideIndex(presentation.slides)
          : (defaultSlideIndex ?? 0);
      const startIndex = clamp(requestedIndex, 0, presentation.slides.length - 1);
      const startSlide = presentation.slides[startIndex];

      // Capture raw XML before materialization clears it (lazy mode).
      // Used for priority font detection: first-slide typefaces decode first.
      const startSlideXml = startSlide?.sourceXml;

      // The active slide's nodes must be reliable for subscribers the moment
      // the store reports "ready", even in lazy mode.
      if (startSlide) materializeSlideNodes(presentation, startSlide);

      // Start decoding ALL embedded fonts while still in "loading" state.
      // Priority typefaces (first slide + theme fonts) decode first in the
      // worker queue, but we wait for complete — not just ready — so no
      // embedded font can swap in after slides are visible (no FOUT).
      // Presentations with no embedded fonts resolve instantly (noop handle).
      // Workers run in parallel, so wall time ≈ time of the slowest single font.
      // Font modules are dynamically imported so the decode pipeline (EOT
      // parsing, MTX decompression, worker pool) stays out of the bundle
      // and off the wire when fonts are disabled or the deck embeds none.
      if (options?.embedFonts !== false && presentation.embeddedFonts?.length) {
        const { collectPriorityTypefaces, injectEmbeddedFonts } =
          await import("@diceui/pptx-parser/fonts");
        if (gen !== loadGeneration) throw ABORT_ERROR;
        const priorityTypefaces = collectPriorityTypefaces(presentation, [startSlideXml]);
        fontInjection = injectEmbeddedFonts(presentation, { priorityTypefaces });
        setState({ progress: 95 });
        await fontInjection.complete;
        if (gen !== loadGeneration) throw ABORT_ERROR;
      }

      replaceState({
        status: "ready",
        presentation,
        activeSlideId: startSlide?.id ?? null,
        zoom: 1,
        progress: 100,
        error: null,
      });

      return presentation;
    } catch (err) {
      if (err instanceof DOMException && err.name === "AbortError") throw err;
      if (gen !== loadGeneration) throw ABORT_ERROR;
      const error = err instanceof Error ? err : new Error(String(err));
      setState({ status: "error", error });
      throw error;
    }
  }

  function getSlideIndex(slideId: string): number {
    return slideIndexById.get(slideId) ?? -1;
  }

  function getActiveSlideIndex(): number {
    const { activeSlideId } = state;
    if (!activeSlideId) return -1;
    return slideIndexById.get(activeSlideId) ?? -1;
  }

  function getActiveSlide(): PresentationData["slides"][number] | null {
    const { presentation } = state;
    if (!presentation) return null;
    const index = getActiveSlideIndex();
    return presentation.slides[index] ?? null;
  }

  function goTo(slideId: string): void {
    const { presentation } = state;
    if (!presentation) return;
    if (state.activeSlideId === slideId) return;
    const index = slideIndexById.get(slideId);
    if (index === undefined) return;

    // Keep the active slide's nodes reliable in lazy mode: materialize the
    // target before subscribers observe the navigation. No-op when already
    // materialized.
    const target = presentation.slides[index];
    if (target) materializeSlideNodes(presentation, target);

    setState({ activeSlideId: slideId });
  }

  function goToIndex(index: number): void {
    const { presentation } = state;
    if (!presentation || presentation.slides.length === 0) return;
    const clamped = clamp(index, 0, presentation.slides.length - 1);
    const slideId = presentation.slides[clamped]?.id;
    if (slideId) goTo(slideId);
  }

  function next(): void {
    const index = getActiveSlideIndex();
    if (index === -1) return;
    goToIndex(index + 1);
  }

  function prev(): void {
    const index = getActiveSlideIndex();
    if (index === -1) return;
    goToIndex(index - 1);
  }

  function canGoNext(): boolean {
    const { presentation } = state;
    if (!presentation) return false;
    const index = getActiveSlideIndex();
    return index >= 0 && index < presentation.slides.length - 1;
  }

  function canGoPrev(): boolean {
    return getActiveSlideIndex() > 0;
  }

  function setZoom(zoom: number): void {
    if (!Number.isFinite(zoom)) return;
    const clamped = clamp(zoom, 0.1, 4);
    if (Object.is(clamped, state.zoom)) return;
    setState({ zoom: clamped });
  }

  function zoomIn(step = 0.25): void {
    setZoom(state.zoom + step);
  }

  function zoomOut(step = 0.25): void {
    setZoom(state.zoom - step);
  }

  function fitTo(containerWidth: number, containerHeight: number, padding = 0): void {
    const { presentation } = state;
    if (!presentation) return;
    const availW = containerWidth - padding * 2;
    const availH = containerHeight - padding * 2;
    if (availW <= 0 || availH <= 0) return;
    setZoom(Math.min(availW / presentation.width, availH / presentation.height));
  }

  function reset(): void {
    loadGeneration += 1;
    fontInjection?.dispose();
    fontInjection = undefined;
    replaceState({ ...INITIAL_STATE });
  }

  return {
    getState,
    subscribe,
    load,
    reset,
    goTo,
    goToIndex,
    next,
    prev,
    setZoom,
    zoomIn,
    zoomOut,
    fitTo,
    getSlideIndex,
    getActiveSlideIndex,
    getActiveSlide,
    canGoNext,
    canGoPrev,
  };
}
