import type { PresentationData } from "@diceui/pptx-parser";
import { buildPresentation, parseZip } from "@diceui/pptx-parser";

export type PreviewInput = ArrayBuffer | Uint8Array | Blob | File;
export type PresentationStatus = "idle" | "loading" | "ready" | "error";

export interface PresentationState {
  status: PresentationStatus;
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
  zoom: number;
  progress: number;
  error: Error | null;
}

export interface PresentationStore {
  getState: () => PresentationState;
  subscribe: (listener: () => void) => () => void;

  /**
   * Parse and display a presentation file.
   * Returns the parsed `PresentationData` on success.
   * Rejects with an `AbortError` if a newer `load()` call superseded this one
   * before it completed — callers can safely ignore that error.
   */
  load: (
    input: PreviewInput,
    options?: { defaultSlideIndex?: number },
  ) => Promise<PresentationData>;
  reset: () => void;

  goTo: (slideId: string) => void;
  goToIndex: (index: number) => void;
  next: () => void;
  prev: () => void;

  setZoom: (zoom: number) => void;
  zoomIn: (step?: number) => void;
  zoomOut: (step?: number) => void;
  fitTo: (containerWidth: number, containerHeight: number, padding?: number) => void;

  getActiveSlideIndex: () => number;
  getActiveSlide: () => PresentationData["slides"][number] | null;
  canGoNext: () => boolean;
  canGoPrev: () => boolean;
}

const INITIAL_STATE: PresentationState = {
  status: "idle",
  presentation: null,
  activeSlideId: null,
  zoom: 1,
  progress: 0,
  error: null,
};

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
    options?: { defaultSlideIndex?: number },
  ): Promise<PresentationData> {
    loadGeneration += 1;
    const gen = loadGeneration;

    replaceState({ ...INITIAL_STATE, status: "loading", progress: 0 });

    try {
      const buffer = await normalizeInput(input);
      if (gen !== loadGeneration)
        throw new DOMException("Superseded by a newer load", "AbortError");
      setState({ progress: 30 });

      const files = await parseZip(buffer);
      if (gen !== loadGeneration)
        throw new DOMException("Superseded by a newer load", "AbortError");
      setState({ progress: 70 });

      const presentation = buildPresentation(files);
      if (gen !== loadGeneration)
        throw new DOMException("Superseded by a newer load", "AbortError");

      const startIndex = clamp(options?.defaultSlideIndex ?? 0, 0, presentation.slides.length - 1);
      replaceState({
        status: "ready",
        presentation,
        activeSlideId: presentation.slides[startIndex]?.id ?? null,
        zoom: 1,
        progress: 100,
        error: null,
      });

      return presentation;
    } catch (err) {
      if (err instanceof DOMException && err.name === "AbortError") throw err;
      if (gen !== loadGeneration)
        throw new DOMException("Superseded by a newer load", "AbortError");
      const error = err instanceof Error ? err : new Error(String(err));
      setState({ status: "error", error });
      throw error;
    }
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
    if (!state.presentation) return;
    if (state.activeSlideId === slideId) return;
    if (!slideIndexById.has(slideId)) return;
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
    getActiveSlideIndex,
    getActiveSlide,
    canGoNext,
    canGoPrev,
  };
}
