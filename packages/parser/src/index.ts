// New primary exports (v2 API)
export { PptxViewer } from "./api/pptx-viewer";
export type {
  ViewerOptions,
  FitMode,
  ListRenderOptions,
  ThumbnailRenderOptions,
  SearchHighlightHandle,
  SearchHighlightOptions,
  PptxViewerEventMap,
  PreviewInput,
} from "./api/pptx-viewer";

export { parseZip, parseZipLazyMedia, RECOMMENDED_ZIP_LIMITS } from "./ooxml/zip-parser";
export type { ZipParseLimits } from "./ooxml/zip-parser";
export type { MediaResolver, ResolvedMedia } from "./media/resolve";

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

export { serializePresentation } from "./serialize/presentation";
export type {
  SerializedPresentation,
  SerializedSlide,
  SerializedNode,
} from "./serialize/presentation";

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
export type { PdfjsOptions, PdfjsConfig } from "./media/pdf-renderer";

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
export type { PicNodeData, CropRect } from "./model/nodes/picture-node";
export type { TableNodeData, TableCell, TableRow } from "./model/nodes/table-node";
export type { GroupNodeData } from "./model/nodes/group-node";
export type { ChartNodeData } from "./model/nodes/chart-node";
export type { PptxFiles } from "./ooxml/zip-parser";

// Font utilities
export { deobfuscateFont } from "./fonts/font-deobfuscate";
export { collectPriorityTypefaces, injectEmbeddedFonts } from "./fonts/font-injector";
export type { FontInjectionHandle, InjectEmbeddedFontsOptions } from "./fonts/font-injector";
