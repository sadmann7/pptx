export * as Presentation from "./primitive";

export type { ErrorProps, ErrorState } from "./error";
export type { LoadingProps, LoadingState } from "./loading";
export type { RootProps, RootState } from "./root";
export type { SelectionChangeEvent, SelectionProps, SelectionState } from "./selection";
export type { SlideProps, SlideState } from "./slide";
export type {
  ThumbnailSelectEvent,
  ThumbnailItemNumberProps,
  ThumbnailItemPreviewProps,
  ThumbnailItemPreviewState,
  ThumbnailItemProps,
  ThumbnailItemState,
  ThumbnailListProps,
  ThumbnailListRenderState,
  ThumbnailListState,
} from "./thumbnail-list";
export type { ViewportProps, ViewportState } from "./viewport";

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
  EditEvent,
  EditSource,
  HistoryChangeEvent,
  StoreState as PresentationState,
  PresentationStatus,
  Store as PresentationStore,
  StoreEventMap as PresentationStoreEventMap,
  PreviewInput,
  SidePadding,
  SlideChangeEvent,
  SlideChangeReason,
  StatusChangeEvent,
  ZoomChangeEvent,
  ZoomChangeReason,
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
