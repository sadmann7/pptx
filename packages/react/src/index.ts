import { Root } from "./components/root";
import { Viewport } from "./components/viewport";
import { Slide } from "./components/slide";
import { Thumbnails } from "./components/thumbnails";

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
