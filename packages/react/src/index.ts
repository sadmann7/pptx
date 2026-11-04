export * as Presentation from "./primitives";

export type { ErrorProps, ErrorState } from "./error";
export type { LoadingProps, LoadingState } from "./loading";
export type { RootProps, RootState } from "./root";
export type { SelectionChangeEvent, SelectionProps, SelectionState } from "./selection/selection";
export type { MarqueeProps } from "./selection/marquee";
export type { ResizeHandleProps, ResizeHandleState } from "./selection/resize-handle";
export type { RotateHandleProps, RotateHandleState } from "./selection/rotate-handle";
export type { SelectionBoxProps, SelectionBoxState } from "./selection/selection-box";
export type { SnapGuideProps, SnapGuideState } from "./selection/snap-guide";
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

export { KEEP_SELECTION_ATTRIBUTE, useSelection } from "./selection/channel";
export type { DrawTool, UseSelectionResult } from "./selection/channel";
export type { ActiveTextStyle } from "./selection/text-range";

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
  AddImageOperation,
  AddShapeOperation,
  AddTextBoxOperation,
  AddSlideOperation,
  BatchOperation,
  DeleteNodeOperation,
  DeleteSlideOperation,
  DuplicateNodeOperation,
  DuplicateSlideOperation,
  EditOperation,
  EditResult,
  LayoutData,
  MoveSlideOperation,
  NodeClipboard,
  NodeOrder,
  ParagraphAlign,
  ParagraphStyle,
  PasteNodesOperation,
  PresentationData,
  PresentationFonts,
  ReorderNodeOperation,
  SetNodeTransformOperation,
  SetOutlineOperation,
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

export { copyNodes, getPresentationFonts, getPresetShapePath } from "@diceui/pptx-core";
