export * as Presentation from "./presentation";

// Flat type re-exports — mirrors Base UI's select/index.ts pattern.
// Allows: import type { RootProps, ViewportState } from '@diceui/pptx'
export type * from "./components/root";
export type * from "./components/viewport";
export type * from "./components/slide";
export type * from "./components/loading";
export type * from "./components/error";
export type * from "./components/thumbnail-list";

export { usePresentation, useSlide, useZoom } from "./context";
export type { UsePresentationResult, UseSlideResult, UseZoomResult } from "./context";

export type { RenderProp, ComponentRenderFn } from "./utils/render";

export type {
  PresentationState,
  PresentationStore,
  PresentationStatus,
  PreviewInput,
} from "./store";
