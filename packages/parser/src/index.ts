// ─── Main API ─────────────────────────────────────────────────────────────────
export { parsePresentation } from "./parse";

// ─── Types ───────────────────────────────────────────────────────────────────
export type {
  // Root
  Presentation,
  Slide,
  SlideSize,
  ParseOptions,
  PresentationInput,

  // Elements
  SlideElement,
  BaseElement,
  GeometricShape,
  ImageShape,
  TextShape,
  TableShape,
  ChartShape,
  GroupShape,
  ConnectorShape,

  // Text
  Paragraph,
  ParagraphContent,
  ParagraphStyle,
  TextRun,
  LineBreak,
  FieldRun,
  RunStyle,
  BodyProperties,
  Bullet,
  BulletNone,
  BulletAuto,
  BulletChar,
  BulletNumeric,
  LineSpacing,
  TextAlignment,
  TextVerticalAlignment,
  TextDirection,
  UnderlineStyle,

  // Table
  TableRow,
  TableCell,

  // Fill & stroke
  Fill,
  SolidFill,
  GradientFill,
  GradientStop,
  PatternFill,
  NoFill,
  Stroke,
  ArrowEnd,
  ArrowEndType,
  ArrowEndSize,

  // Color
  Color,
  SolidColor,
  SchemeColor,

  // Theme
  Theme,
  ThemeColors,
  ThemeFonts,

  // Geometry
  Position,
  Size,
  Transform,
  Background,

  // Effects
  Effect,
  OuterShadow,

  // Placeholders
  PlaceholderInfo,
  PlaceholderType,
} from "./types";

// ─── Color utilities ──────────────────────────────────────────────────────────
export { resolveColor } from "./color";

// ─── EMU utilities ────────────────────────────────────────────────────────────
export { emuToPoints, hunPtToPoints, perMilleToPercent, angleToDegs } from "./emu";
