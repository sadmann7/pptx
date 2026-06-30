import { Root } from "./components/root";
import { Viewport } from "./components/viewport";
import { Slide } from "./components/slide";
import {
  ThumbnailList,
  ThumbnailItem,
  ThumbnailItemPreview,
  ThumbnailItemNumber,
} from "./components/thumbnail-list";
import { Loading } from "./components/loading";
import { PresentationError as Error } from "./components/error";

export const Presentation = {
  Root,
  Viewport,
  Slide,
  ThumbnailList,
  ThumbnailItem,
  ThumbnailItemPreview,
  ThumbnailItemNumber,
  Loading,
  Error,
} as const;

export { usePresentation, useSlide, useZoom } from "./context";
export type { UsePresentationResult, UseSlideResult, UseZoomResult } from "./context";

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
  ThumbnailItemPreviewState,
  ThumbnailItemPreviewProps,
  ThumbnailItemNumberProps,
} from "./components/thumbnail-list";
export type { LoadingState, LoadingProps } from "./components/loading";
export type { ErrorState, PresentationErrorProps } from "./components/error";
