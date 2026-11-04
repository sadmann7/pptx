export * as Presentation from "./primitives";

export type { ErrorProps, ErrorState } from "./error";
export type { LoadingProps, LoadingState } from "./loading";
export type { RootProps, RootState } from "./root";
export type { SelectionChangeEvent, SelectionProps, SelectionState } from "./selection";
export type { MarqueeProps } from "./marquee";
export type { ResizeHandleProps, ResizeHandleState } from "./resize-handle";
export type { RotateHandleProps, RotateHandleState } from "./rotate-handle";
export type { SelectionBoxProps, SelectionBoxState } from "./selection-box";
export type { SnapGuideProps, SnapGuideState } from "./snap-guide";
export type { SlideProps, SlideState } from "./slide";
export type {
  ThumbnailItemNumberProps,
  ThumbnailItemPreviewProps,
  ThumbnailItemPreviewState,
  ThumbnailItemProps,
  ThumbnailItemState,
  ThumbnailListProps,
  ThumbnailListRenderState,
  ThumbnailListState,
  ThumbnailSelectEvent,
} from "./thumbnail-list";
export type { ViewportProps, ViewportState } from "./viewport";

export {
  useCreateStore as useCreatePresentationStore,
  useHistory,
  usePresentation,
  useStore as usePresentationStore,
  useSlide,
  useSlideIndex,
  useSlideRevision,
  useZoom,
} from "./context";
export type {
  UseHistoryResult,
  UsePresentationResult,
  UseSlideResult,
  UseZoomResult,
} from "./context";

export type { ComponentRenderFn, RenderProp } from "./render";

export { KEEP_SELECTION_ATTRIBUTE, useSelection } from "./selection-channel";
export type { UseSelectionResult } from "./selection-channel";
export type { ActiveTextStyle } from "./text-range";

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
  AddSlideOperation,
  BatchOperation,
  DeleteNodeOperation,
  DeleteSlideOperation,
  DuplicateSlideOperation,
  EditOperation,
  EditResult,
  LayoutData,
  MoveSlideOperation,
  ParagraphAlign,
  ParagraphStyle,
  PresentationData,
  SetNodeTransformOperation,
  SetParagraphStyleOperation,
  SetSlideBackgroundOperation,
  SetSolidFillOperation,
  SetTextBodyOperation,
  SetTextBodyParagraph,
  SetTextBodyRun,
  SetTextRunOperation,
  SetTextStyleOperation,
  TextPosition,
  TextRange,
  TextStyle,
} from "@diceui/pptx-core";
