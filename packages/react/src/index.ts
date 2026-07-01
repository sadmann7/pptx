export * as Presentation from "./primitive";

export type * from "./root";
export type * from "./viewport";
export type * from "./slide";
export type * from "./loading";
export type * from "./error";
export type * from "./thumbnail-list";

export { useCreatePresentationStore, usePresentation, useSlide, useZoom } from "./context";
export type { UsePresentationResult, UseSlideResult, UseZoomResult } from "./context";

export type { RenderProp, ComponentRenderFn } from "./render";

export type {
  PresentationState,
  PresentationStore,
  PresentationStatus,
  PreviewInput,
} from "./store";
