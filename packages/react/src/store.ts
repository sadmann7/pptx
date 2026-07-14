import type {
  EditOperation,
  EditResult,
  FontInjectionHandle,
  PptxSaveOptions,
  PresentationData,
  SlideData,
} from "@diceui/pptx-core";
import {
  applyEdit,
  buildPresentation,
  materializeSlide,
  readPptx,
  writePptx,
} from "@diceui/pptx-core";

import { DEFAULT_STORE_STATE } from "./constant";

export const MIN_ZOOM = 0.1;
export const MAX_ZOOM = 4;
export const DEFAULT_ZOOM_STEP = 0.25;

const ABORT_ERROR = new DOMException("Superseded by a newer load", "AbortError");

// Load-progress cost model. Work is measured in byte-equivalent units where
// 1 unit ≈ unzipping one compressed input byte; these ratios express how
// expensive the other phases are per byte relative to that baseline.
const READ_COST_PER_BYTE = 0.05;
const FONT_DECODE_COST_PER_BYTE = 20;

export interface SidePadding {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export type AutoFitPadding = number | Partial<SidePadding>;

interface HistoryEntry {
  op: EditOperation;
  result: EditResult;
}

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
export interface StoreState {
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

  /**
   * Monotonic counter bumped on every applied edit, undo, or redo.
   *
   * The `presentation` object is mutated in place by edits, so its identity
   * never changes; subscribe to `revision` to react to content changes.
   * Resets to `0` on `load()` and `reset()`.
   */
  revision: number;
}

/** Store returned by `useCreatePresentationStore` for controlled usage. */
export interface Store {
  /** Returns the current state snapshot. Non-reactive: use `subscribe` to watch for changes. */
  getState: () => StoreState;
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
   * } catch (error) {
   *   if (error instanceof DOMException && error.name === "AbortError") return;
   *   console.error(error);
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

      /**
       * When `false`, the source package is retained so the presentation can
       * be edited (`store.edit()`) and saved back to a .pptx (`store.save()`).
       * The retained package keeps the source zip in memory (compressed).
       *
       * Follows the same convention as `<input readOnly>`: omitting the prop
       * (or passing `true`) gives a read-only viewer; pass `false` to enable
       * editing.
       *
       * @default true
       */
      readOnly?: boolean;
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
   * Clamped to `[MIN_ZOOM, MAX_ZOOM]`.
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
  fitTo: (containerWidth: number, containerHeight: number, padding?: AutoFitPadding) => void;

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

  /**
   * Apply an edit operation to the loaded presentation.
   *
   * Requires the deck to have been loaded with `{ readOnly: false }`.
   * On success the edit is pushed onto the undo stack, the redo stack is
   * cleared, and affected slides get a new revision so mounted views
   * re-render.
   *
   * ```ts
   * await store.edit({
   *   type: "setTextRun",
   *   slideId, nodeId: "2",
   *   paragraphIndex: 0, runIndex: 0,
   *   text: "Hello",
   * });
   * ```
   */
  edit: (op: EditOperation) => Promise<EditResult>;

  /** Revert the most recent edit. Returns `false` when the undo stack is empty. */
  undo: () => boolean;

  /** Re-apply the most recently undone edit. Resolves `false` when the redo stack is empty. */
  redo: () => Promise<boolean>;

  /** Returns `true` when there is an edit to undo. */
  canUndo: () => boolean;

  /** Returns `true` when there is an edit to redo. */
  canRedo: () => boolean;

  /**
   * Serialize the (possibly edited) presentation back to a .pptx archive.
   * Requires the deck to have been loaded with `{ readOnly: false }`.
   */
  save: (options?: PptxSaveOptions) => Promise<Uint8Array>;

  /**
   * Returns the render revision of a slide: bumped whenever an edit, undo,
   * or redo affects that slide. Safe to call inside `useSyncExternalStore`
   * selectors.
   */
  getSlideRevision: (slideId: string) => number;

  /**
   * Returns the content revision of a slide: like `getSlideRevision`, but
   * NOT bumped by transform-only edits (`setNodeTransform` moves/resizes).
   * Views can compare this against a previous value to know whether a
   * revision bump requires a full re-render or just a position patch.
   */
  getSlideContentRevision: (slideId: string) => number;
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

function normalizePadding(padding: AutoFitPadding | undefined = 0): SidePadding {
  if (typeof padding === "number") {
    return { top: padding, right: padding, bottom: padding, left: padding };
  }

  return {
    top: padding?.top ?? 0,
    right: padding?.right ?? 0,
    bottom: padding?.bottom ?? 0,
    left: padding?.left ?? 0,
  };
}

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}

export function createStore(): Store {
  let state: StoreState = { ...DEFAULT_STORE_STATE };
  const listeners = new Set<() => void>();
  let loadGeneration = 0;
  let fontInjection: FontInjectionHandle | undefined;

  let slideIndexById = new Map<string, number>();

  let undoStack: HistoryEntry[] = [];
  let redoStack: HistoryEntry[] = [];
  let slideRevisionById = new Map<string, number>();
  let slideContentRevisionById = new Map<string, number>();

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
    next: Partial<StoreState> | ((current: StoreState) => Partial<StoreState>),
  ): void {
    const patch = typeof next === "function" ? next(state) : next;
    state = { ...state, ...patch };
    if ("presentation" in patch) rebuildSlideIndex(state.presentation);
    emit();
  }

  function replaceState(next: StoreState): void {
    state = next;
    rebuildSlideIndex(next.presentation);
    emit();
  }

  function getState(): StoreState {
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
      readOnly?: boolean;
    },
  ): Promise<PresentationData> {
    loadGeneration += 1;
    const gen = loadGeneration;

    fontInjection?.dispose();
    fontInjection = undefined;
    clearEditHistory();
    replaceState({ ...DEFAULT_STORE_STATE, status: "loading", progress: 0 });

    // Progress = workDone / workTotal in byte-equivalent units (see cost
    // model above). The budget grows once font bytes are known after unzip;
    // the monotonic clamp keeps the bar from moving backward when it does.
    const inputBytes = Math.max(1, input instanceof Blob ? input.size : input.byteLength);
    const readUnits = inputBytes * READ_COST_PER_BYTE;
    const zipUnits = inputBytes;

    let workDone = 0;
    let workTotal = readUnits + zipUnits;
    let reportedProgress = 0;
    const reportProgress = (): void => {
      if (gen !== loadGeneration) return;
      // Hold at 99 while loading; only the ready state reports 100.
      const pct = Math.min(99, Math.floor((workDone / workTotal) * 100));
      if (pct > reportedProgress) {
        reportedProgress = pct;
        setState({ progress: pct });
      }
    };

    try {
      const buffer = await normalizeInput(input);
      if (gen !== loadGeneration) throw ABORT_ERROR;
      workDone = readUnits;
      reportProgress();

      const files = await readPptx(buffer, {
        lazyMedia: true,
        keepPackage: options?.readOnly === false,
        onProgress: (done, total) => {
          workDone = readUnits + zipUnits * (done / total);
          reportProgress();
        },
      });
      if (gen !== loadGeneration) throw ABORT_ERROR;
      workDone = readUnits + zipUnits;
      reportProgress();

      // Lazy by default: per-slide node parsing is deferred until a slide
      // is rendered, keeping load time flat for large decks.
      const presentation = buildPresentation(files, { lazy: options?.lazy ?? true });
      if (gen !== loadGeneration) throw ABORT_ERROR;

      const defaultSlideIndex = options?.defaultSlideIndex;
      const requestedIndex =
        typeof defaultSlideIndex === "function"
          ? defaultSlideIndex(presentation.slides)
          : (defaultSlideIndex ?? 0);
      const startIndex = clamp(requestedIndex, 0, presentation.slides.length - 1);
      const startSlide = presentation.slides[startIndex];

      // Capture raw XML for priority font detection before materialization
      // clears it (lazy mode).
      const startSlideXml = startSlide?.sourceXml;

      // The active slide's nodes must be reliable the moment the store
      // reports "ready", even in lazy mode.
      if (startSlide) materializeSlide(presentation, startSlide);

      // Decode embedded fonts while still "loading" and wait for all of them,
      // so no font can swap in after slides are visible (no FOUT). The font
      // pipeline is a lazily imported chunk; decks without embedded fonts
      // (or embedFonts: false) never fetch it.
      if (options?.embedFonts !== false && presentation.embeddedFonts?.length) {
        const { collectPriorityTypefaces, injectEmbeddedFonts } =
          await import("@diceui/pptx-core/fonts");
        if (gen !== loadGeneration) throw ABORT_ERROR;

        // Expand the progress budget by the now-known decode workload.
        // Dedupe: the fonts map can alias one part under two paths.
        let fontBytes = 0;
        for (const part of new Set(presentation.fonts.values())) {
          fontBytes += part.byteLength;
        }
        const fontUnits = fontBytes * FONT_DECODE_COST_PER_BYTE;
        const fontBaseUnits = workDone;
        workTotal += fontUnits;

        const priorityTypefaces = collectPriorityTypefaces(presentation, [startSlideXml]);
        fontInjection = injectEmbeddedFonts(presentation, {
          priorityTypefaces,
          onProgress: (done, total) => {
            workDone = fontBaseUnits + fontUnits * (done / total);
            reportProgress();
          },
        });
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
        revision: 0,
      });

      return presentation;
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") throw error;
      if (gen !== loadGeneration) throw ABORT_ERROR;
      const typedError = error instanceof Error ? error : new Error(String(error));
      setState({ status: "error", error: typedError });
      throw typedError;
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

    // Materialize the target before subscribers observe the navigation
    // (lazy mode); no-op when already materialized.
    const target = presentation.slides[index];
    if (target) materializeSlide(presentation, target);

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
    const clamped = clamp(zoom, MIN_ZOOM, MAX_ZOOM);
    if (Object.is(clamped, state.zoom)) return;
    setState({ zoom: clamped });
  }

  function zoomIn(step = DEFAULT_ZOOM_STEP): void {
    setZoom(state.zoom + step);
  }

  function zoomOut(step = DEFAULT_ZOOM_STEP): void {
    setZoom(state.zoom - step);
  }

  function fitTo(
    containerWidth: number,
    containerHeight: number,
    padding: AutoFitPadding = 0,
  ): void {
    const { presentation } = state;
    if (!presentation) return;
    const { top, right, bottom, left } = normalizePadding(padding);
    const availW = containerWidth - left - right;
    const availH = containerHeight - top - bottom;
    if (availW <= 0 || availH <= 0) return;
    setZoom(Math.min(availW / presentation.width, availH / presentation.height));
  }

  function reset(): void {
    loadGeneration += 1;
    fontInjection?.dispose();
    fontInjection = undefined;
    clearEditHistory();
    replaceState({ ...DEFAULT_STORE_STATE });
  }

  function clearEditHistory(): void {
    undoStack = [];
    redoStack = [];
    slideRevisionById = new Map();
    slideContentRevisionById = new Map();
  }

  function getSlideRevision(slideId: string): number {
    return slideRevisionById.get(slideId) ?? 0;
  }

  function getSlideContentRevision(slideId: string): number {
    return slideContentRevisionById.get(slideId) ?? 0;
  }

  /** `true` when the operation only changes node transforms (position/size/rotation). */
  function isTransformOnly(op: EditOperation): boolean {
    if (op.type === "setNodeTransform") return true;
    if (op.type === "batch") return op.operations.every(isTransformOnly);
    return false;
  }

  /**
   * Publish the outcome of an edit/undo/redo: bump affected slide revisions,
   * repair navigation if the active slide disappeared, and notify
   * subscribers. The presentation object is mutated in place, so the state
   * `revision` counter is the only identity change subscribers see.
   */
  function commitEdit(
    affectedSlideIds: string[],
    prevActiveIndex: number,
    navigateToFirstAffected = false,
    transformOnly = false,
  ): void {
    const { presentation } = state;
    if (!presentation) return;

    for (const slideId of affectedSlideIds) {
      slideRevisionById.set(slideId, getSlideRevision(slideId) + 1);
      if (!transformOnly) {
        slideContentRevisionById.set(slideId, getSlideContentRevision(slideId) + 1);
      }
    }
    rebuildSlideIndex(presentation);

    let activeSlideId = state.activeSlideId;
    const activeGone = activeSlideId !== null && !slideIndexById.has(activeSlideId);
    if (activeGone || activeSlideId === null) {
      const fallbackIndex = clamp(prevActiveIndex, 0, presentation.slides.length - 1);
      activeSlideId = presentation.slides[fallbackIndex]?.id ?? null;
    }
    // Undo/redo: jump to the slide the action touched so the user sees the
    // change (PowerPoint behavior). Only when that slide still exists.
    if (navigateToFirstAffected) {
      const target = affectedSlideIds.find((id) => slideIndexById.has(id));
      if (target) activeSlideId = target;
    }
    const activeSlide = activeSlideId
      ? presentation.slides[slideIndexById.get(activeSlideId) ?? -1]
      : undefined;
    if (activeSlide) materializeSlide(presentation, activeSlide);

    setState({ activeSlideId, revision: state.revision + 1 });
  }

  async function edit(op: EditOperation): Promise<EditResult> {
    const { presentation, status } = state;
    if (!presentation || status !== "ready") {
      throw new Error("PresentationStore.edit: no presentation is loaded");
    }
    if (!presentation.sourcePackage) {
      throw new Error(
        "PresentationStore.edit: presentation was loaded read-only; pass { readOnly: false } to load()",
      );
    }

    const prevActiveIndex = getActiveSlideIndex();
    const result = await applyEdit(presentation, op);
    undoStack.push({ op, result });
    redoStack = [];
    commitEdit(result.affectedSlideIds, prevActiveIndex, false, isTransformOnly(op));
    return result;
  }

  function undo(): boolean {
    const entry = undoStack.pop();
    if (!entry) return false;
    const prevActiveIndex = getActiveSlideIndex();
    entry.result.undo();
    redoStack.push(entry);
    // Undo of a duplicate removes the created slide; include it so any
    // mounted view of it is invalidated.
    const affected = entry.result.createdSlideId
      ? [...entry.result.affectedSlideIds, entry.result.createdSlideId]
      : entry.result.affectedSlideIds;
    commitEdit(affected, prevActiveIndex, true, isTransformOnly(entry.op));
    return true;
  }

  async function redo(): Promise<boolean> {
    const entry = redoStack.pop();
    if (!entry) return false;
    const { presentation } = state;
    if (!presentation) return false;

    const prevActiveIndex = getActiveSlideIndex();
    // Re-apply the original operation; it captures fresh undo state.
    const result = await applyEdit(presentation, entry.op);
    undoStack.push({ op: entry.op, result });
    commitEdit(result.affectedSlideIds, prevActiveIndex, true, isTransformOnly(entry.op));
    return true;
  }

  function canUndo(): boolean {
    return undoStack.length > 0;
  }

  function canRedo(): boolean {
    return redoStack.length > 0;
  }

  async function save(options?: PptxSaveOptions): Promise<Uint8Array> {
    const { presentation } = state;
    if (!presentation) {
      throw new Error("PresentationStore.save: no presentation is loaded");
    }
    if (!presentation.sourcePackage) {
      throw new Error(
        "PresentationStore.save: presentation was loaded read-only; pass { readOnly: false } to load()",
      );
    }
    return writePptx(presentation, options);
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
    edit,
    undo,
    redo,
    canUndo,
    canRedo,
    save,
    getSlideRevision,
    getSlideContentRevision,
  };
}
