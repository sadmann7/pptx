import type { ParseOptions, Presentation, PresentationInput } from "@pptx/parser";
import { parsePresentation } from "@pptx/parser";

// ─── State ───────────────────────────────────────────────────────────────────

export type PresentationStatus = "idle" | "loading" | "ready" | "error";

export interface PresentationState {
  status: PresentationStatus;
  presentation: Presentation | null;
  /** 0-based current slide index */
  currentIndex: number;
  /** Scale factor. 1 = native pt dimensions. Use fit() to auto-compute. */
  zoom: number;
  /** Bytes loaded during parsing, 0–100 */
  progress: number;
  error: Error | null;
}

const INITIAL_STATE: PresentationState = {
  status: "idle",
  presentation: null,
  currentIndex: 0,
  zoom: 1,
  progress: 0,
  error: null,
};

// ─── Store ───────────────────────────────────────────────────────────────────

/**
 * Plain-object store compatible with React.useSyncExternalStore.
 *
 * One store instance lives per <Presentation.Root>. All compound components
 * read from it via context + useSyncExternalStore selectors.
 */
export class PresentationStore {
  private state: PresentationState = { ...INITIAL_STATE };
  private listeners = new Set<() => void>();
  /** Track load calls so stale responses from old files are discarded */
  private loadGeneration = 0;

  // ── useSyncExternalStore interface ──────────────────────────────────────

  getState(): PresentationState {
    return this.state;
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  // ── Actions ─────────────────────────────────────────────────────────────

  async load(input: PresentationInput, options?: Omit<ParseOptions, "onProgress">): Promise<void> {
    this.loadGeneration++;
    const gen = this.loadGeneration;

    this.setState({
      ...INITIAL_STATE,
      status: "loading",
      progress: 0,
    });

    try {
      const presentation = await parsePresentation(input, {
        ...options,
        onProgress: (current, total) => {
          if (gen !== this.loadGeneration) return;
          this.setState({ ...this.state, progress: Math.round((current / total) * 100) });
        },
      });

      if (gen !== this.loadGeneration) return;

      this.setState({
        status: "ready",
        presentation,
        currentIndex: 0,
        zoom: 1,
        progress: 100,
        error: null,
      });
    } catch (err) {
      if (gen !== this.loadGeneration) return;
      this.setState({
        ...this.state,
        status: "error",
        error: err instanceof Error ? err : new Error(String(err)),
      });
    }
  }

  goTo(index: number): void {
    if (!this.state.presentation) return;
    const clamped = Math.max(0, Math.min(this.state.presentation.slides.length - 1, index));
    if (clamped === this.state.currentIndex) return;
    this.setState({ ...this.state, currentIndex: clamped });
  }

  next(): void {
    this.goTo(this.state.currentIndex + 1);
  }

  prev(): void {
    this.goTo(this.state.currentIndex - 1);
  }

  setZoom(zoom: number): void {
    const clamped = Math.max(0.1, Math.min(4, zoom));
    if (clamped === this.state.zoom) return;
    this.setState({ ...this.state, zoom: clamped });
  }

  zoomIn(step = 0.25): void {
    this.setZoom(this.state.zoom + step);
  }

  zoomOut(step = 0.25): void {
    this.setZoom(this.state.zoom - step);
  }

  /**
   * Compute the zoom needed to fill `containerWidth × containerHeight`
   * while maintaining the slide aspect ratio.
   *
   * containerWidth/Height are in CSS pixels (from getBoundingClientRect / clientWidth).
   * slideSize is in CSS points. 1pt = 96/72px in browsers, so we must convert.
   */
  fitTo(containerWidth: number, containerHeight: number, padding = 0): void {
    if (!this.state.presentation) return;
    const { width, height } = this.state.presentation.slideSize;
    const availW = containerWidth - padding * 2;
    const availH = containerHeight - padding * 2;
    // Convert slide dimensions from pt to px before computing the ratio
    const PT_TO_PX = 96 / 72;
    const zoom = Math.min(availW / (width * PT_TO_PX), availH / (height * PT_TO_PX));
    this.setZoom(zoom);
  }

  reset(): void {
    this.loadGeneration++;
    this.setState({ ...INITIAL_STATE });
  }

  // ── Internal ─────────────────────────────────────────────────────────────

  private setState(state: PresentationState): void {
    this.state = state;
    for (const listener of this.listeners) listener();
  }
}
