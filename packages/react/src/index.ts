export * as Presentation from "./primitive";

export type * from "./error";
export type * from "./loading";
export type * from "./root";
export type * from "./selection";
export type * from "./slide";
export type * from "./thumbnail-list";
export type * from "./viewport";

export {
  useCreateStore as useCreatePresentationStore,
  usePresentation,
  useSlide,
  useSlideIndex,
  useSlideRevision,
  useZoom,
} from "./context";
export type { UsePresentationResult, UseSlideResult, UseZoomResult } from "./context";

export type { ComponentRenderFn, RenderProp } from "./render";

export type {
  AutoFitPadding,
  StoreState as PresentationState,
  PresentationStatus,
  Store as PresentationStore,
  PreviewInput,
  SidePadding,
} from "./store";

export type {
  BatchOperation,
  DeleteNodeOperation,
  DeleteSlideOperation,
  DuplicateSlideOperation,
  EditOperation,
  EditResult,
  MoveSlideOperation,
  PresentationData,
  SetNodeTransformOperation,
  SetSolidFillOperation,
  SetTextBodyOperation,
  SetTextBodyParagraph,
  SetTextBodyRun,
  SetTextRunOperation,
} from "@diceui/pptx-core";
