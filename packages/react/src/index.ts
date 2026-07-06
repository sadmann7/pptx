export * as Presentation from "./primitive";

export type * from "./error";
export type * from "./loading";
export type * from "./root";
export type * from "./slide";
export type * from "./thumbnail-list";
export type * from "./viewport";

export { useCreatePresentationStore, usePresentation, useSlide, useZoom } from "./context";
export type { UsePresentationResult, UseSlideResult, UseZoomResult } from "./context";

export type { ComponentRenderFn, RenderProp } from "./render";

export type {
  PresentationState,
  PresentationStatus,
  PresentationStore,
  PreviewInput,
} from "./store";

export type {
  DeleteNodeOperation,
  DeleteSlideOperation,
  DuplicateSlideOperation,
  EditOperation,
  EditResult,
  MoveSlideOperation,
  PresentationData,
  SetNodeTransformOperation,
  SetSolidFillOperation,
  SetTextRunOperation,
} from "@diceui/pptx-parser";
