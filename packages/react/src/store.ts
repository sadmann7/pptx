import { parseZip, buildPresentation, materializeSlideNodes } from "@pptx/parser";
import type { PresentationData } from "@pptx/parser";

export type PreviewInput = ArrayBuffer | Uint8Array | Blob | File;
export type PresentationStatus = "idle" | "loading" | "ready" | "error";

export interface PresentationState {
  status: PresentationStatus;
  presentation: PresentationData | null;
  currentIndex: number;
  zoom: number;
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

async function normalizeInput(input: PreviewInput): Promise<ArrayBuffer> {
  if (input instanceof ArrayBuffer) return input;
  if (input instanceof Uint8Array)
    return input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength) as ArrayBuffer;
  if (input instanceof Blob) return input.arrayBuffer();
  throw new Error("Unsupported input type");
}

export class PresentationStore {
  private state: PresentationState = { ...INITIAL_STATE };
  private listeners = new Set<() => void>();
  private loadGeneration = 0;

  getState(): PresentationState {
    return this.state;
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  async load(input: PreviewInput): Promise<void> {
    this.loadGeneration++;
    const gen = this.loadGeneration;
    this.setState({ ...INITIAL_STATE, status: "loading", progress: 0 });

    try {
      const buffer = await normalizeInput(input);
      if (gen !== this.loadGeneration) return;
      this.setState({ ...this.state, progress: 30 });

      const files = await parseZip(buffer);
      if (gen !== this.loadGeneration) return;
      this.setState({ ...this.state, progress: 70 });

      const presentation = buildPresentation(files);
      if (gen !== this.loadGeneration) return;

      // Materialize all slides eagerly so they're ready to render
      for (const slide of presentation.slides) {
        materializeSlideNodes(presentation, slide);
      }

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

  fitTo(containerWidth: number, containerHeight: number, padding = 0): void {
    if (!this.state.presentation) return;
    const { width, height } = this.state.presentation;
    const availW = containerWidth - padding * 2;
    const availH = containerHeight - padding * 2;
    const zoom = Math.min(availW / width, availH / height);
    this.setZoom(zoom);
  }

  reset(): void {
    this.loadGeneration++;
    this.setState({ ...INITIAL_STATE });
  }

  private setState(state: PresentationState): void {
    this.state = state;
    for (const listener of this.listeners) listener();
  }
}
