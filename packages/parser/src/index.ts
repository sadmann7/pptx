// New primary exports (v2 API)
export { PptxViewer } from "./core/viewer";
export type {
  ViewerOptions,
  FitMode,
  ListRenderOptions,
  ThumbnailRenderOptions,
  SearchHighlightHandle,
  SearchHighlightOptions,
  PptxViewerEventMap,
  PreviewInput,
} from "./core/viewer";

// Deprecated aliases (v1 compat)
export { PptxRenderer } from "./core/renderer";
export type { RendererOptions } from "./core/renderer";

export { parseZip, parseZipLazyMedia, RECOMMENDED_ZIP_LIMITS } from "./parser/zip-parser";
export type { ZipParseLimits } from "./parser/zip-parser";
export type { MediaResolver, ResolvedMedia } from "./utils/media";

export {
  buildPresentation,
  materializeAllSlideNodes,
  materializeSlideNodes,
} from "./model/presentation";
export type { BuildPresentationOptions, PresentationData } from "./model/presentation";

export { serializePresentation } from "./export/serialize-presentation";
export type {
  SerializedPresentation,
  SerializedSlide,
  SerializedNode,
} from "./export/serialize-presentation";

// Model-level text search
export { buildTextIndex, searchPresentation, searchText } from "./search/text-search";
export type {
  SearchTextKind,
  TextBounds,
  TextIndexEntry,
  TextIndexOptions,
  TextSearchOptions,
  TextSearchResult,
} from "./search/text-search";

// Headless single-slide rendering
export { renderSlide } from "./renderer/slide-renderer";
export type { SlideHandle, SlideRendererOptions } from "./renderer/slide-renderer";
export type { PdfjsOptions, PdfjsConfig } from "./utils/pdf-renderer";

// Model types
export type { SlideData, SlideNode } from "./model/slide";
export type { ThemeData } from "./model/theme";
export type {
  BaseNodeData,
  Position,
  Size,
  NodeType,
  PlaceholderInfo,
  HlinkAction,
} from "./model/nodes/base-node";
export type {
  ShapeNodeData,
  TextBody,
  TextParagraph,
  TextRun,
  LineEndInfo,
  TextBoxBounds,
} from "./model/nodes/shape-node";
export type { PicNodeData, CropRect } from "./model/nodes/pic-node";
export type { TableNodeData, TableCell, TableRow } from "./model/nodes/table-node";
export type { GroupNodeData } from "./model/nodes/group-node";
export type { ChartNodeData } from "./model/nodes/chart-node";
export type { PptxFiles } from "./parser/zip-parser";
