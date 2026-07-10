// New primary exports (v2 API)
export { PptxViewer } from "./api/pptx-viewer";
export type {
  FitMode,
  ListRenderOptions,
  PptxViewerEventMap,
  PreviewInput,
  SearchHighlightHandle,
  SearchHighlightOptions,
  ThumbnailRenderOptions,
  ViewerOptions,
} from "./api/pptx-viewer";

export { applyEdit } from "./edit/operations";
export type {
  BatchOperation,
  DeleteNodeOperation,
  DeleteSlideOperation,
  DuplicateSlideOperation,
  EditOperation,
  EditResult,
  MoveSlideOperation,
  SetNodeTransformOperation,
  SetSolidFillOperation,
  SetTextBodyOperation,
  SetTextBodyParagraph,
  SetTextBodyRun,
  SetTextRunOperation,
} from "./edit/operations";
export type { MediaResolver, ResolvedMedia } from "./media/resolve";
export { PptxPackage } from "./ooxml/package";
export type { PptxSaveOptions } from "./ooxml/package";
export { writePptx } from "./ooxml/pptx-writer";
export { parseZip, parseZipLazyMedia, RECOMMENDED_ZIP_LIMITS } from "./ooxml/zip-parser";
export type { ZipParseLimits, ZipParseOptions } from "./ooxml/zip-parser";

export {
  buildPresentation,
  materializeAllSlideNodes,
  materializeSlideNodes,
} from "./model/presentation";
export type {
  BuildPresentationOptions,
  EmbeddedFontEntry,
  EmbeddedFontVariant,
  PresentationData,
} from "./model/presentation";

export { serializePresentation } from "./model/serialize";
export type { SerializedNode, SerializedPresentation, SerializedSlide } from "./model/serialize";

// Model-level text search
export { buildTextIndex, searchPresentation, searchText } from "./model/text-search";
export type {
  SearchTextKind,
  TextBounds,
  TextIndexEntry,
  TextIndexOptions,
  TextSearchOptions,
  TextSearchResult,
} from "./model/text-search";

// Headless single-slide rendering
export type { PdfjsConfig, PdfjsOptions } from "./media/pdf-renderer";
export { renderSlide, renderThumbnail } from "./renderer/slide-renderer";
export type {
  SlideHandle,
  SlideRendererOptions,
  ThumbnailRendererOptions,
} from "./renderer/slide-renderer";

// Model types
export type {
  BaseNodeData,
  HlinkAction,
  NodeType,
  PlaceholderInfo,
  Position,
  Size,
} from "./model/nodes/base-node";
export type { ChartNodeData } from "./model/nodes/chart-node";
export type { GroupNodeData } from "./model/nodes/group-node";
export type { CropRect, PicNodeData } from "./model/nodes/picture-node";
export type {
  LineEndInfo,
  ShapeNodeData,
  TextBody,
  TextBoxBounds,
  TextParagraph,
  TextRun,
} from "./model/nodes/shape-node";
export type { TableCell, TableNodeData, TableRow } from "./model/nodes/table-node";
export type { SlideData, SlideNode } from "./model/slide";
export type { ThemeData } from "./model/theme";
export type { PptxFiles } from "./ooxml/zip-parser";

// Embedded font support lives in a separate entry ("@diceui/pptx-parser/fonts")
// so the decode pipeline is only loaded when actually used. Only the handle
// type is re-exported here for consumers typing against the API.
export type { FontInjectionHandle, InjectEmbeddedFontsOptions } from "./fonts/font-injector";
