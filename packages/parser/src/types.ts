// ─── Coordinates ─────────────────────────────────────────────────────────────
// All measurements are in points (pt) after conversion from EMU.
// 1 pt = 12700 EMU. This makes values directly usable as CSS px at 96dpi.

export interface Position {
  x: number;
  y: number;
}

export interface Size {
  width: number;
  height: number;
}

export interface SlideSize {
  width: number;
  height: number;
}

// ─── Colors ──────────────────────────────────────────────────────────────────

export interface SolidColor {
  type: "solid";
  /** 6-character hex string, no leading '#' */
  hex: string;
  /** 0–100 */
  alpha: number;
}

export interface SchemeColor {
  type: "scheme";
  /** e.g. 'accent1', 'dk1', 'lt1', 'tx1', 'bg1', 'hlink' */
  token: string;
  lumMod?: number;
  lumOff?: number;
  shade?: number;
  tint?: number;
  alpha?: number;
}

export type Color = SolidColor | SchemeColor;

// ─── Theme ───────────────────────────────────────────────────────────────────

export interface ThemeColors {
  dk1: string;
  dk2: string;
  lt1: string;
  lt2: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  hlink: string;
  folHlink: string;
}

export interface ThemeFonts {
  major: string;
  minor: string;
}

export interface Theme {
  name: string;
  colors: ThemeColors;
  fonts: ThemeFonts;
}

// ─── Text ────────────────────────────────────────────────────────────────────

export type TextAlignment = "left" | "center" | "right" | "justify" | "dist" | "thai";
export type TextVerticalAlignment = "top" | "mid" | "bottom" | "just" | "dist";
export type TextDirection = "horz" | "vert" | "vert270" | "wordArtVert";
export type UnderlineStyle =
  | "sng"
  | "dbl"
  | "thick"
  | "dash"
  | "dotDash"
  | "dotDotDash"
  | "wavy"
  | "none";

export interface BulletNone {
  type: "none";
}

export interface BulletAuto {
  type: "auto";
  char?: string;
  color?: Color;
  size?: number;
  fontFamily?: string;
}

export interface BulletChar {
  type: "char";
  char: string;
  color?: Color;
  size?: number;
  fontFamily?: string;
}

export interface BulletNumeric {
  type: "numeric";
  style:
    | "arabicPeriod"
    | "arabicParenR"
    | "romanLcPeriod"
    | "romanUcPeriod"
    | "alphaLcParenR"
    | "alphaUcPeriod"
    | string;
  startAt?: number;
  color?: Color;
  size?: number;
}

export type Bullet = BulletNone | BulletAuto | BulletChar | BulletNumeric;

export interface LineSpacing {
  /** Points if 'pt', percent if 'pct' */
  value: number;
  unit: "pt" | "pct";
}

export interface ParagraphStyle {
  alignment?: TextAlignment;
  /** Indent level 0–8 for lists */
  level?: number;
  indent?: number;
  marginLeft?: number;
  marginRight?: number;
  spaceBefore?: LineSpacing;
  spaceAfter?: LineSpacing;
  lineSpacing?: LineSpacing;
  bullet?: Bullet;
  defaultRunStyle?: RunStyle;
}

export interface RunStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: UnderlineStyle;
  strikethrough?: boolean;
  /** Font size in points */
  fontSize?: number;
  color?: Color;
  fontFamily?: string;
  /** 'major' | 'minor' resolves via Theme.fonts */
  fontTheme?: "major" | "minor";
  highlight?: Color;
  baseline?: number;
  link?: string;
  language?: string;
  noProof?: boolean;
  dirty?: boolean;
}

export interface TextRun {
  type: "run";
  text: string;
  style: RunStyle;
}

export interface LineBreak {
  type: "lineBreak";
}

export interface FieldRun {
  type: "field";
  /** e.g. 'slidenum', 'datetime' */
  fieldType: string;
  text: string;
  style: RunStyle;
}

export type ParagraphContent = TextRun | LineBreak | FieldRun;

export interface Paragraph {
  runs: ParagraphContent[];
  style: ParagraphStyle;
}

export interface BodyProperties {
  verticalAlignment?: TextVerticalAlignment;
  direction?: TextDirection;
  wrap?: "none" | "square";
  autofit?: "none" | "spAutoFit" | "normAutoFit";
  insetLeft?: number;
  insetRight?: number;
  insetTop?: number;
  insetBottom?: number;
  columns?: number;
  columnSpacing?: number;
  rotation?: number;
}

// ─── Fill & Stroke ───────────────────────────────────────────────────────────

export interface SolidFill {
  type: "solid";
  color: Color;
}

export interface GradientStop {
  position: number;
  color: Color;
}

export interface GradientFill {
  type: "gradient";
  stops: GradientStop[];
  angle?: number;
}

export interface PatternFill {
  type: "pattern";
  preset: string;
  fgColor?: Color;
  bgColor?: Color;
}

export interface NoFill {
  type: "none";
}

export type Fill = SolidFill | GradientFill | PatternFill | NoFill;

export type ArrowEndType = "none" | "triangle" | "stealth" | "diamond" | "oval" | "arrow";
export type ArrowEndSize = "sm" | "med" | "lg";

export interface ArrowEnd {
  type: ArrowEndType;
  /** Width of the arrowhead relative to the line width */
  width?: ArrowEndSize;
  /** Length of the arrowhead */
  length?: ArrowEndSize;
}

export interface Stroke {
  fill: Fill;
  /** Width in points */
  width?: number;
  dashStyle?: "solid" | "dot" | "dash" | "dashDot" | "lgDash" | "lgDashDot" | string;
  joinStyle?: "round" | "bevel" | "miter";
  capStyle?: "flat" | "sq" | "rnd";
  /** Arrowhead at the start of the line */
  headEnd?: ArrowEnd;
  /** Arrowhead at the end of the line */
  tailEnd?: ArrowEnd;
}

// ─── Transform ───────────────────────────────────────────────────────────────

export interface Transform {
  /** Clockwise rotation in degrees */
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
}

// ─── Shadow & Effects ────────────────────────────────────────────────────────

export interface OuterShadow {
  type: "outer";
  color: Color;
  /** In points */
  blurRadius?: number;
  distance?: number;
  direction?: number;
  alignment?: string;
}

export type Effect = OuterShadow;

// ─── Shapes ──────────────────────────────────────────────────────────────────

export interface BaseElement {
  id: string;
  name?: string;
  position: Position;
  size: Size;
  transform?: Transform;
  hidden?: boolean;
  /**
   * Set when this element is a placeholder slot. The renderer can use this
   * to apply additional placeholder-specific logic (e.g. "click to edit title").
   */
  placeholder?: PlaceholderInfo;
}

/** Preset or custom geometry shape (sp) */
export interface GeometricShape extends BaseElement {
  type: "shape";
  /**
   * OOXML preset shape name e.g. 'rect', 'ellipse', 'roundRect',
   * 'rightArrow', 'star5', etc.  'custom' for custom geometry.
   */
  shapeType: string;
  fill?: Fill;
  stroke?: Stroke;
  effects?: Effect[];
  /** Text body if the shape contains text */
  body?: {
    paragraphs: Paragraph[];
    properties: BodyProperties;
  };
}

/** Picture (pic) */
export interface ImageShape extends BaseElement {
  type: "image";
  /** Blob URL or data URI populated after ZIP extraction */
  src: string;
  mimeType: string;
  /** Original embedded path for reference */
  rId: string;
  cropRect?: {
    top: number;
    right: number;
    bottom: number;
    left: number;
  };
  fill?: Fill;
  stroke?: Stroke;
  effects?: Effect[];
}

/** Text box (sp with no preset geometry) */
export interface TextShape extends BaseElement {
  type: "text";
  paragraphs: Paragraph[];
  properties: BodyProperties;
  fill?: Fill;
  stroke?: Stroke;
}

/** Table (graphicFrame > tbl) */
export interface TableCell {
  rowSpan: number;
  colSpan: number;
  paragraphs: Paragraph[];
  fill?: Fill;
  stroke?: Stroke;
  /** True if this cell is a continuation of a merged cell */
  merged: boolean;
}

export interface TableRow {
  /** Height in points */
  height?: number;
  cells: TableCell[];
}

export interface TableShape extends BaseElement {
  type: "table";
  /** Column widths in points */
  columnWidths: number[];
  rows: TableRow[];
}

/** Chart (graphicFrame > chart) — opaque, renderer can use chartData */
export interface ChartShape extends BaseElement {
  type: "chart";
  /** rId of the chart relationship */
  rId: string;
  /** Raw chart XML string — let the renderer decide what to do with it */
  chartXml: string;
}

/** Group shape (grpSp) */
export interface GroupShape extends BaseElement {
  type: "group";
  children: SlideElement[];
}

/** Connector shape (cxnSp) */
export interface ConnectorShape extends BaseElement {
  type: "connector";
  shapeType: string;
  fill?: Fill;
  stroke?: Stroke;
}

export type SlideElement =
  | GeometricShape
  | ImageShape
  | TextShape
  | TableShape
  | ChartShape
  | GroupShape
  | ConnectorShape;

// ─── Placeholder ─────────────────────────────────────────────────────────────

/**
 * OOXML placeholder types.
 * Placeholders are named slots on a layout/master that slides inherit from.
 */
export type PlaceholderType =
  | "title"
  | "ctrTitle"
  | "subTitle"
  | "body"
  | "obj"
  | "pic"
  | "tbl"
  | "chart"
  | "dgm"
  | "media"
  | "dt"
  | "sldNum"
  | "ftr"
  | "hdr"
  | string;

export interface PlaceholderInfo {
  type: PlaceholderType;
  /** Index within the layout, used to match multiple body placeholders */
  idx: number;
}

// ─── Background ──────────────────────────────────────────────────────────────

export interface Background {
  fill: Fill;
}

// ─── Slide ───────────────────────────────────────────────────────────────────

export interface Slide {
  /** 0-based index */
  index: number;
  /** Relationship ID from presentation.xml */
  rId: string;
  /** The embedded path e.g. 'ppt/slides/slide1.xml' */
  path: string;
  elements: SlideElement[];
  background?: Background;
  notes?: string;
  /** Whether slide is hidden in the presentation */
  hidden?: boolean;
}

// ─── Presentation ────────────────────────────────────────────────────────────

export interface Presentation {
  slideSize: SlideSize;
  theme: Theme;
  slides: Slide[];
}

// ─── Parser input ────────────────────────────────────────────────────────────

export type PresentationInput = ArrayBuffer | Uint8Array | Blob | string; // base64 or URL — callers should fetch and pass ArrayBuffer

export interface ParseOptions {
  /**
   * Called for each slide index as it is parsed. Useful for showing
   * incremental progress in the UI.
   */
  onProgress?: (current: number, total: number) => void;
  /**
   * If true, skip notes extraction (faster for thumbnail use cases).
   * @default false
   */
  skipNotes?: boolean;
  /**
   * If true, do not extract images — ImageShape.src will be an empty string.
   * Useful for metadata-only parsing.
   * @default false
   */
  skipImages?: boolean;
}
