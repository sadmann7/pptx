import { Root } from "./components/Root";
import { Viewport } from "./components/Viewport";
import { Slide } from "./components/Slide";
import { Thumbnails } from "./components/Thumbnails";

export const Presentation = {
  Root,
  Viewport,
  Slide,
  Thumbnails,
} as const;

export { usePresentation, useSlide, useZoom, usePresentationStoreRef } from "./context";
export type { UsePresentationResult, UseSlideResult, UseZoomResult } from "./context";
export { PresentationStore } from "./store";
export type { PresentationState, PresentationStatus, PreviewInput } from "./store";
