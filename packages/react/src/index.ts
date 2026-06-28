import { Root } from "./components/root";
import { Viewport } from "./components/viewport";
import { Slide } from "./components/slide";
import { ThumbnailList, ThumbnailItem, ThumbnailItemCanvas } from "./components/thumbnails";
import { Loading } from "./components/loading";
import { PresentationError as Error } from "./components/error";

export const Presentation = {
  Root,
  Viewport,
  Slide,
  ThumbnailList,
  ThumbnailItem,
  ThumbnailItemCanvas,
  Loading,
  Error,
} as const;

export { usePresentation, useSlide, useZoom, usePresentationStoreRef } from "./context";
export type { UsePresentationResult, UseSlideResult, UseZoomResult } from "./context";
export { PresentationStore } from "./store";
export type { PresentationState, PresentationStatus, PreviewInput } from "./store";

// Render-prop utilities — exposed so consumers can build custom components
// that compose cleanly with the rest of the system.
export { mergeProps, mergeRefs, renderElement } from "./utils/render";
export type { RenderProp, ComponentRenderFn } from "./utils/render";

// Component state types
export type { ViewportState, ViewportProps } from "./components/viewport";
export type { SlideState, SlideProps } from "./components/slide";
export type {
  ThumbnailListState,
  ThumbnailListRenderState,
  ThumbnailListProps,
  ThumbnailItemState,
  ThumbnailItemProps,
  ThumbnailItemCanvasState,
  ThumbnailItemCanvasProps,
} from "./components/thumbnails";
export type { LoadingState, LoadingProps } from "./components/loading";
export type { ErrorState, PresentationErrorProps } from "./components/error";
