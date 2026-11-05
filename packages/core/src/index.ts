export { applySlideScale, PptxViewer } from "./viewer";
export type {
  FitMode,
  ListRenderOptions,
  PptxSource,
  PptxViewerEventMap,
  PptxViewerOptions,
  SearchHighlightHandle,
  SearchHighlightOptions,
  ThumbnailRenderOptions,
} from "./viewer";

export { applyEdit, copyNodes } from "./edit/operation";
export type {
  AddImageOperation,
  AddShapeOperation,
  AddTextBoxOperation,
  AddSlideOperation,
  AddTableOperation,
  BatchOperation,
  ClipboardRelationship,
  DeleteNodeOperation,
  DeleteSlideOperation,
  DuplicateNodeOperation,
  DuplicateSlideOperation,
  EditOperation,
  EditResult,
  MoveSlideOperation,
  NodeClipboard,
  NodeOrder,
  ParagraphAlign,
  ParagraphStyle,
  PasteNodesOperation,
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
  TableCellAddress,
  TextPosition,
  TextRange,
  TextStyle,
} from "./edit/operation";
export type { MediaResolver, ResolvedMedia } from "./media/resolve";
export { PptxPackage } from "./ooxml/package";
export type { PptxSaveOptions } from "./ooxml/package";
export { writePptx } from "./ooxml/writer";
export { readPptx, RECOMMENDED_PPTX_READ_LIMITS } from "./ooxml/zip";
export type { PptxReadLimits, PptxReadOptions } from "./ooxml/zip";

export { buildPresentation, materializeAllSlides, materializeSlide } from "./model/presentation";
export type {
  BuildPresentationOptions,
  EmbeddedFontEntry,
  EmbeddedFontVariant,
  PresentationData,
} from "./model/presentation";
export type { LayoutData } from "./model/layout";
export { getPresetShapePath } from "./geometry/preset";
export { getPresentationFonts } from "./model/presentation-fonts";
export type { PresentationFonts } from "./model/presentation-fonts";

export { serializePresentation } from "./model/serialize";
export type { SerializedNode, SerializedPresentation, SerializedSlide } from "./model/serialize";

export { buildTextIndex, searchPresentation, searchText } from "./model/text-search";
export type {
  SearchTextKind,
  TextBounds,
  TextIndexEntry,
  TextIndexOptions,
  TextSearchOptions,
  TextSearchResult,
} from "./model/text-search";

export type { PdfjsConfig, PdfjsOptions } from "./media/pdf";
export { PPTX_ATTRS, PPTX_DATASET } from "./renderer/dom-attributes";
export { renderSlide, renderThumbnail } from "./renderer/slide";
export type { SlideHandle, SlideRendererOptions, ThumbnailRendererOptions } from "./renderer/slide";

export type {
  BaseNodeData,
  HyperlinkAction,
  NodePosition,
  NodeSize,
  NodeType,
  PlaceholderInfo,
} from "./model/nodes/base";
export type { ChartNodeData } from "./model/nodes/chart";
export type { GroupNodeData } from "./model/nodes/group";
export type { CropRect, PictureNodeData } from "./model/nodes/picture";
export type {
  LineEndInfo,
  ShapeNodeData,
  TextBody,
  TextBoxBounds,
  TextParagraph,
  TextRun,
} from "./model/nodes/shape";
export type { TableCell, TableNodeData, TableRow } from "./model/nodes/table";
export type { SlideData, SlideNode } from "./model/slide";
export type { ThemeData } from "./model/theme";
export type { PptxFiles } from "./ooxml/zip";
