// Parser
export { parseXml, SafeXmlNode } from "./parser/XmlParser";
export { parseZip, parseZipLazyMedia, RECOMMENDED_ZIP_LIMITS } from "./parser/ZipParser";
export type { PptxFiles, ZipParseLimits } from "./parser/ZipParser";
export { parseRels, resolveRelTarget, isExternalTargetMode } from "./parser/RelParser";
export type { RelEntry } from "./parser/RelParser";
export { emuToPx, emuToPt, angleToDeg, pctToDecimal, ptToPx } from "./parser/units";
export { parseOoxmlBool } from "./parser/booleans";

// Model
export {
  buildPresentation,
  materializeAllSlideNodes,
  materializeSlideNodes,
} from "./model/Presentation";
export type { BuildPresentationOptions, PresentationData } from "./model/Presentation";
export { parseSlide, createLazySlide, materializeSlideData } from "./model/Slide";
export type { SlideData, SlideNode } from "./model/Slide";
export { parseTheme } from "./model/Theme";
export type { ThemeData, ThemeFontInfo } from "./model/Theme";
export type { RenderableNode } from "./model/RenderableChild";
export { isPlaceholderNode, parseRenderableChild } from "./model/RenderableChild";

// Nodes
export { parseBaseProps } from "./model/nodes/BaseNode";
export type {
  BaseNodeData,
  Position,
  Size,
  NodeType,
  PlaceholderInfo,
  HlinkAction,
} from "./model/nodes/BaseNode";
export { parseShapeNode, parseTextBody } from "./model/nodes/ShapeNode";
export type {
  ShapeNodeData,
  TextBody,
  TextParagraph,
  TextRun,
  LineEndInfo,
  TextBoxBounds,
} from "./model/nodes/ShapeNode";
export { parsePicNode } from "./model/nodes/PicNode";
export type { PicNodeData, CropRect } from "./model/nodes/PicNode";
export { parseTableNode } from "./model/nodes/TableNode";
export type { TableNodeData, TableCell, TableRow } from "./model/nodes/TableNode";
export { parseGroupNode } from "./model/nodes/GroupNode";
export type { GroupNodeData } from "./model/nodes/GroupNode";
export { parseChartNode } from "./model/nodes/ChartNode";
export type { ChartNodeData } from "./model/nodes/ChartNode";

// Utils
export {
  hexToRgb,
  rgbToHex,
  hslToRgb,
  rgbToHsl,
  applyTint,
  applyShade,
  applyColorModifiers,
  presetColorToHex,
} from "./utils/color";
export type { ColorModifier } from "./utils/color";
export {
  getMimeType,
  resolveMediaPath,
  findMediaByTarget,
  findMediaByTargetAsync,
  getOrCreateBlobUrl,
} from "./utils/media";
export type { MediaResolver, ResolvedMedia } from "./utils/media";
export { isAllowedExternalUrl, isAllowedExternalMediaUrl } from "./utils/urlSafety";

// Input types
export type PreviewInput = ArrayBuffer | Uint8Array | Blob;
