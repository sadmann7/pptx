# OOXML (ECMA-376) Support Tracker

Feature-level support status of `@diceui/pptx-core` against **ECMA-376 5th
edition** (PresentationML + DrawingML). The spec is freely available from
[Ecma International](https://ecma-international.org/publications-and-standards/standards/ecma-376/)
(Part 1 PDF plus the DrawingML geometry and schema archives).

Legend: ✅ supported · 🟡 partial · ❌ not supported

Keep this file in sync when adding or discovering spec behavior, and every ✅
claim should be backed by tests in `src/tests/` (see `.cursor/rules/testing.mdc`).

## Package structure (Part 2, OPC)

| Feature                                                         | Status | Notes                                                        |
| --------------------------------------------------------------- | ------ | ------------------------------------------------------------ |
| ZIP container + `[Content_Types].xml`                           | ✅     | `ooxml/zip.ts`; read limits (`RECOMMENDED_PPTX_READ_LIMITS`) |
| Relationships (`.rels`, internal/external, URI-encoded targets) | ✅     | `ooxml/rel-parser.ts`                                        |
| Lazy media loading                                              | ✅     | `readPptx({ lazyMedia: true })` defers media inflation       |
| Presentation / slides / layouts / masters / themes parts        | ✅     | Full inheritance chain slide → layout → master → theme       |
| Notes slides / notes masters                                    | ❌     | Parts are ignored                                            |
| Handout masters, comments, custom shows                         | ❌     |                                                              |

## Slides & inheritance (Part 1 §19)

| Feature                                                                    | Status | Notes                              |
| -------------------------------------------------------------------------- | ------ | ---------------------------------- |
| Slide size (`sldSz`), slide order (`sldIdLst`), hidden slides              | ✅     |                                    |
| Placeholder inheritance (position/size/bodyPr/lstStyle from layout/master) | ✅     |                                    |
| Color map + `clrMapOvr`                                                    | ✅     |                                    |
| Backgrounds: `bgPr` (solid/gradient/pattern/picture) & `bgRef`             | ✅     | Slide > layout > master precedence |
| `showMasterSp`                                                             | ✅     |                                    |
| Slide transitions (`p:transition`)                                         | ❌     | Static viewer                      |
| Animations/timing (`p:timing`)                                             | ❌     | Static viewer                      |

## Shapes & geometry (DrawingML §20.1)

| Feature                                                             | Status | Notes                                                                                                                  |
| ------------------------------------------------------------------- | ------ | ---------------------------------------------------------------------------------------------------------------------- |
| All 187 preset geometries (`ST_ShapeType`)                          | ✅     | `geometry/presets.ts`; hand-coded, verified against `presetShapeDefinitions.xml`; multi-path presets for shaded shapes |
| Adjustment values (`avLst`) on presets                              | ✅     | Defaults audited against the spec                                                                                      |
| Custom geometry (`custGeom`): all 6 path commands                   | ✅     | `moveTo`, `lnTo`, `cubicBezTo`, `quadBezTo`, `arcTo`, `close`                                                          |
| `custGeom` guide formulas (`gdLst`, all 17 ops + built-ins)         | ✅     | `geometry/guide-evaluator.ts`                                                                                          |
| `custGeom` per-path `fill`/`stroke` attributes                      | ❌     | Sub-paths render as one merged path                                                                                    |
| Adjust handles (`ahLst`), connection sites (`cxnLst`)               | ❌     | Relevant for future editing                                                                                            |
| Preset text warp (`prstTxWarp`)                                     | 🟡     | Recognized; only basic WordArt vertical handled, no warp geometry                                                      |
| Transforms: rotation, flipH/flipV                                   | ✅     |                                                                                                                        |
| Groups (`grpSp`): child coordinate space, nesting, flips, `grpFill` | ✅     |                                                                                                                        |
| Connectors (`cxnSp`) + head/tail ends (arrow markers)               | ✅     | Endpoint routing only; no shape-attached rerouting                                                                     |

## Fills, lines, effects (DrawingML §20.1.8)

| Feature                                                        | Status | Notes                                                        |
| -------------------------------------------------------------- | ------ | ------------------------------------------------------------ |
| `solidFill` (srgbClr/schemeClr/sysClr/prstClr/hslClr/scrgbClr) | ✅     |                                                              |
| Color modifiers (tint, shade, lumMod/Off, satMod, alpha, …)    | ✅     | Full `applyColorModifiers` set incl. linear-space tint/shade |
| `gradFill`: linear, radial (`path="circle"`), rect             | ✅     |                                                              |
| `pattFill` (preset patterns)                                   | ✅     | SVG patterns                                                 |
| `blipFill` on shapes and backgrounds (stretch, tile, srcRect)  | ✅     |                                                              |
| `noFill`, `grpFill`                                            | ✅     |                                                              |
| Theme style refs: `fillRef`, `lnRef`, `effectRef` with `phClr` | ✅     |                                                              |
| Outlines: width, dash, caps, joins                             | ✅     |                                                              |
| `outerShdw`                                                    | ✅     | SVG `feDropShadow` / CSS drop-shadow                         |
| `glow`, `softEdge`, `reflection`, `innerShdw`                  | 🟡     | Recognized in places; not faithfully rendered                |
| 3-D (`sp3d`, `scene3d`, bevels)                                | ❌     |                                                              |

## Text (DrawingML §21.1)

| Feature                                                                                | Status | Notes                                                                                                   |
| -------------------------------------------------------------------------------------- | ------ | ------------------------------------------------------------------------------------------------------- |
| Runs: b/i/u/strike, size, color, highlight, spacing, caps, baseline                    | ✅     |                                                                                                         |
| Fonts: explicit `latin`/`ea`/`cs`, theme `+mj-lt`/`+mn-lt`                             | ✅     |                                                                                                         |
| Paragraphs: alignment, RTL, indent, margins, line/para spacing (pct + pts)             | ✅     |                                                                                                         |
| Bullets: `buChar`, `buAutoNum` (all formats), `buClr`, `buSzPct/Pts`, `buFont`, levels | ✅     | Known deviation: `startAt` re-seeds on every paragraph (`TODO(spec?)`)                                  |
| Fields (`a:fld`)                                                                       | 🟡     | Renders cached literal, not computed value (`TODO(spec?)`)                                              |
| Hyperlinks: external URLs, `ppaction://` slide jumps                                   | ✅     | URL protocol allow-list (`utils/url-validation.ts`)                                                     |
| `bodyPr`: anchor, insets, wrap, `normAutofit`, `spAutoFit`, vertical text              | ✅     | Autofit re-measures via DOM; browser metrics ≠ DirectWrite exactly                                      |
| Leading spaces / tabs at line start                                                    | ✅     | `white-space: pre-wrap` applied to runs that start a visual line with spaces; matches PowerPoint layout |
| Embedded fonts (`.fntdata`: EOT/MTX, ODTTF deobfuscation)                              | ✅     | Internal MTX decompressor (`fonts/mtx/`), worker pool, priority loading                                 |
| WordArt / text effects                                                                 | ❌     | Beyond vertical orientation                                                                             |
| Math (OMML)                                                                            | ❌     |                                                                                                         |

## Tables (DrawingML §21.1.3)

| Feature                                                                   | Status | Notes                                              |
| ------------------------------------------------------------------------- | ------ | -------------------------------------------------- |
| Grid, spans, merges; sized from grid (frame `ext` ignored per PowerPoint) | ✅     |                                                    |
| Cell fills, borders, insets, anchors                                      | ✅     |                                                    |
| Table styles: embedded `tblStyleLst` + all 74 built-in styles             | ✅     | firstRow/lastRow/bandRow/firstCol/lastCol/bandCol  |
| Table with no `a:tableStyleId`                                            | ✅     | Falls back to built-in "No Style, Table Grid"      |
| Missing `ppt/tableStyles.xml` part                                        | 🟡     | Built-ins skipped when part absent (`TODO(spec?)`) |

## Charts (DrawingML Charts §21.2)

| Feature                                                                      | Status | Notes                                                           |
| ---------------------------------------------------------------------------- | ------ | --------------------------------------------------------------- |
| bar/column (clustered, stacked, percent), line, area, pie, doughnut, scatter | ✅     | Rendered via echarts                                            |
| Cartesian combo charts + secondary axes                                      | ✅     |                                                                 |
| Axes, legends, titles, manual layouts, data labels, data table               | ✅     |                                                                 |
| Number formats (`formatCode`)                                                | 🟡     | Common codes; plain `0.00` drops trailing zeros (`TODO(spec?)`) |
| Chart styles/colors parts (`style*.xml`, `colors*.xml`)                      | ✅     |                                                                 |
| radar, bubble, stock                                                         | 🟡     | Parsed; simplified rendering                                    |
| surface, 3-D chart variants                                                  | ❌     |                                                                 |

## Media & embedded objects

| Feature                                                          | Status | Notes                                                                    |
| ---------------------------------------------------------------- | ------ | ------------------------------------------------------------------------ |
| Images: PNG/JPEG/GIF/BMP/WebP/SVG, crop (`srcRect`), preset clip | ✅     |                                                                          |
| Duotone/grayscale/biLevel image effects                          | 🟡     | Canvas-based; browser-dependent                                          |
| EMF/WMF                                                          | 🟡     | `media/emf-parser.ts` subset; PDF-in-EMF fallback via pdfjs              |
| Video/audio placeholders                                         | 🟡     | Poster + native controls for supported formats; no OLE                   |
| SmartArt (`dgm`)                                                 | 🟡     | Rendered from `diagrams/drawing*.xml` fallback, not the layout algorithm |
| OLE objects                                                      | ❌     |                                                                          |

## Known spec deviations (tracked as `TODO(spec?)` in tests)

1. `buAutoNum` `startAt` restarts numbering on every paragraph that declares it.
2. `a:fld type="slidenum"` renders the cached literal instead of the computed slide number.
3. `formatValue(2.5, "0.00")` → `"2.5"` (Excel renders `"2.50"`).
4. Built-in table styles are skipped when the package omits `ppt/tableStyles.xml`.
