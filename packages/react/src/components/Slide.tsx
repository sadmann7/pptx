import React from "react";
import { usePresentation, useSlide, useZoom } from "../context";
import { materializeSlideNodes, emuToPx } from "@pptx/parser";
import type { PresentationData, SlideNode, SafeXmlNode } from "@pptx/parser";
import type { ShapeNodeData, TextParagraph } from "@pptx/parser";
import type { PicNodeData } from "@pptx/parser";
import type { GroupNodeData } from "@pptx/parser";
import type { TableNodeData } from "@pptx/parser";
import { getOrCreateBlobUrl, findMediaByTarget } from "@pptx/parser";
import type { RelEntry } from "@pptx/parser";

export interface SlideProps {
  children?: React.ReactNode;
  className?: string;
  style?: React.CSSProperties;
}

export function Slide({ children, className, style }: SlideProps) {
  const { presentation, status } = usePresentation();
  const { slide } = useSlide();
  const { zoom } = useZoom();
  const mediaUrlCache = React.useRef(new Map<string, string>()).current;

  if (status === "loading") return <div className={className} style={{ display: "flex", alignItems: "center", justifyContent: "center", ...style }}><span style={{ color: "#666" }}>Loading...</span></div>;
  if (status === "error") return <div className={className} style={style}>Failed to parse</div>;
  if (!presentation || !slide) return null;

  materializeSlideNodes(presentation, slide);
  const { width, height } = presentation;

  return (
    <div className={className} style={{ display: "flex", alignItems: "center", justifyContent: "center", overflow: "auto", ...style }}>
      <div style={{ width: width * zoom, height: height * zoom, flexShrink: 0 }}>
        <SlideCanvas width={width} height={height} zoom={zoom} nodes={slide.nodes} presentation={presentation} slideRels={slide.rels} mediaUrlCache={mediaUrlCache}>
          {children}
        </SlideCanvas>
      </div>
    </div>
  );
}

interface SlideCanvasProps {
  width: number; height: number; zoom: number;
  nodes: SlideNode[]; presentation: PresentationData;
  slideRels: Map<string, RelEntry>; mediaUrlCache: Map<string, string>;
  children?: React.ReactNode;
}

export function SlideCanvas({ width, height, zoom, nodes, presentation, slideRels, mediaUrlCache, children }: SlideCanvasProps) {
  return (
    <div style={{
      position: "relative", width, height,
      background: "#ffffff", overflow: "hidden",
      transformOrigin: "top left", transform: `scale(${zoom})`,
      color: "#000", fontFamily: "'Calibri', 'Arial', sans-serif",
    }}>
      {nodes.map((node, i) => (
        <NodeRenderer key={node.id || i} node={node} presentation={presentation} slideRels={slideRels} mediaUrlCache={mediaUrlCache} />
      ))}
      {children}
    </div>
  );
}

// ─── Color/Fill resolution from SafeXmlNode ────────────────────────────────

function resolveColorNode(node: SafeXmlNode): string | undefined {
  if (!node.exists()) return undefined;
  const srgb = node.child("srgbClr");
  if (srgb.exists()) {
    const val = srgb.attr("val");
    if (val) return `#${val}`;
  }
  const sysClr = node.child("sysClr");
  if (sysClr.exists()) {
    const lastClr = sysClr.attr("lastClr") ?? sysClr.attr("val");
    if (lastClr) return `#${lastClr}`;
  }
  return undefined;
}

function resolveFillColor(fillNode: SafeXmlNode | undefined): string | undefined {
  if (!fillNode || !fillNode.exists()) return undefined;
  const ln = fillNode.localName;
  if (ln === "solidFill") return resolveColorNode(fillNode);
  if (ln === "noFill") return "transparent";
  if (ln === "gradFill") {
    const stops = fillNode.child("gsLst").children("gs");
    if (stops.length > 0) return resolveColorNode(stops[0]!);
  }
  return undefined;
}

function resolveRunColor(rPr: SafeXmlNode | undefined): string | undefined {
  if (!rPr || !rPr.exists()) return undefined;
  const solidFill = rPr.child("solidFill");
  if (solidFill.exists()) return resolveColorNode(solidFill);
  return undefined;
}

function resolveLineColor(lineNode: SafeXmlNode | undefined): { color: string; width: number } | undefined {
  if (!lineNode || !lineNode.exists()) return undefined;
  if (lineNode.child("noFill").exists()) return undefined;
  const solidFill = lineNode.child("solidFill");
  const color = solidFill.exists() ? resolveColorNode(solidFill) : undefined;
  if (!color) return undefined;
  const w = lineNode.numAttr("w");
  return { color, width: w ? emuToPx(w) : 1 };
}

// ─── Node Renderer ──────────────────────────────────────────────────────────

interface NodeProps {
  node: SlideNode; presentation: PresentationData;
  slideRels: Map<string, RelEntry>; mediaUrlCache: Map<string, string>;
}

function NodeRenderer({ node, presentation, slideRels, mediaUrlCache }: NodeProps) {
  const baseStyle: React.CSSProperties = {
    position: "absolute",
    left: node.position.x, top: node.position.y,
    width: node.size.w, height: node.size.h,
    boxSizing: "border-box",
  };
  const transforms: string[] = [];
  if (node.rotation) transforms.push(`rotate(${node.rotation}deg)`);
  if (node.flipH) transforms.push("scaleX(-1)");
  if (node.flipV) transforms.push("scaleY(-1)");
  if (transforms.length) baseStyle.transform = transforms.join(" ");

  switch (node.nodeType) {
    case "shape": return <ShapeRenderer node={node} style={baseStyle} />;
    case "picture": return <PicRenderer node={node} style={baseStyle} presentation={presentation} slideRels={slideRels} mediaUrlCache={mediaUrlCache} />;
    case "table": return <TableRenderer node={node} style={baseStyle} />;
    case "group": return <GroupRenderer style={baseStyle} />;
    case "chart": return <div style={{ ...baseStyle, border: "1px dashed #ccc", display: "flex", alignItems: "center", justifyContent: "center", color: "#999", fontSize: 12 }}>Chart</div>;
    default: return null;
  }
}

// ─── Shape Renderer ─────────────────────────────────────────────────────────

function ShapeRenderer({ node, style }: { node: ShapeNodeData; style: React.CSSProperties }) {
  const fillColor = resolveFillColor(node.fill);
  const lineStyle = resolveLineColor(node.line);

  const shapeStyle: React.CSSProperties = {
    ...style,
    overflow: "hidden",
    ...(fillColor && fillColor !== "transparent" ? { background: fillColor } : {}),
    ...(lineStyle ? { border: `${lineStyle.width}px solid ${lineStyle.color}` } : {}),
  };

  // Preset geometry → border-radius for common shapes
  if (node.presetGeometry === "ellipse") shapeStyle.borderRadius = "50%";
  else if (node.presetGeometry === "roundRect") shapeStyle.borderRadius = "8px";

  const paragraphs = node.textBody?.paragraphs ?? [];
  const hasText = paragraphs.some(p => p.runs.some(r => r.text.trim()));

  // Body properties for text insets
  const bodyPr = node.textBody?.bodyProperties;
  const lIns = bodyPr?.numAttr("lIns");
  const tIns = bodyPr?.numAttr("tIns");
  const rIns = bodyPr?.numAttr("rIns");
  const bIns = bodyPr?.numAttr("bIns");
  const paddingLeft = lIns !== undefined ? emuToPx(lIns) : 7;
  const paddingTop = tIns !== undefined ? emuToPx(tIns) : 4;
  const paddingRight = rIns !== undefined ? emuToPx(rIns) : 7;
  const paddingBottom = bIns !== undefined ? emuToPx(bIns) : 4;

  // Vertical alignment
  const anchor = bodyPr?.attr("anchor");
  let justifyContent = "flex-start";
  if (anchor === "ctr") justifyContent = "center";
  else if (anchor === "b") justifyContent = "flex-end";

  if (!hasText) return <div style={shapeStyle} />;

  return (
    <div style={{ ...shapeStyle, display: "flex", flexDirection: "column", justifyContent, padding: `${paddingTop}px ${paddingRight}px ${paddingBottom}px ${paddingLeft}px` }}>
      {paragraphs.map((p, pi) => <ParagraphRenderer key={pi} paragraph={p} />)}
    </div>
  );
}

function ParagraphRenderer({ paragraph }: { paragraph: TextParagraph }) {
  const text = paragraph.runs.map(r => r.text).join("");
  if (!text && paragraph.runs.length === 0) return <div style={{ minHeight: "1em" }} />;

  const pPr = paragraph.properties;
  const algn = pPr?.attr("algn");
  let textAlign: React.CSSProperties["textAlign"] = "left";
  if (algn === "ctr") textAlign = "center";
  else if (algn === "r") textAlign = "right";
  else if (algn === "just") textAlign = "justify";

  return (
    <div style={{ textAlign, lineHeight: 1.3 }}>
      {paragraph.runs.map((run, ri) => {
        if (run.text === "\n") return <br key={ri} />;
        if (!run.text) return null;
        const rPr = run.properties;
        const fontSize = rPr?.numAttr("sz");
        const isBold = rPr?.attr("b") === "1";
        const isItalic = rPr?.attr("i") === "1";
        const color = resolveRunColor(rPr);
        const fontFamily = rPr?.child("latin").attr("typeface");

        return (
          <span key={ri} style={{
            fontSize: fontSize ? fontSize / 100 : undefined,
            fontWeight: isBold ? "bold" : undefined,
            fontStyle: isItalic ? "italic" : undefined,
            color: color ?? undefined,
            fontFamily: fontFamily ? `'${fontFamily}', sans-serif` : undefined,
            whiteSpace: "pre-wrap",
          }}>
            {run.text}
          </span>
        );
      })}
    </div>
  );
}

// ─── Picture Renderer ───────────────────────────────────────────────────────

function PicRenderer({ node, style, presentation, slideRels, mediaUrlCache }: { node: PicNodeData; style: React.CSSProperties; presentation: PresentationData; slideRels: Map<string, RelEntry>; mediaUrlCache: Map<string, string> }) {
  const src = React.useMemo(() => {
    const embedId = node.blipEmbed;
    if (!embedId) return undefined;
    const rel = slideRels.get(embedId);
    if (!rel) return undefined;
    const resolved = findMediaByTarget(rel.target, presentation.media);
    if (!resolved) return undefined;
    return getOrCreateBlobUrl(resolved.mediaPath, resolved.data, mediaUrlCache);
  }, [node.blipEmbed, slideRels, presentation.media, mediaUrlCache]);

  if (!src) return <div style={{ ...style, background: "#f5f5f5" }} />;

  const imgStyle: React.CSSProperties = { width: "100%", height: "100%", objectFit: "fill", display: "block" };
  if (node.crop) {
    const { top, right, bottom, left } = node.crop;
    const visibleW = 1 - left - right;
    const visibleH = 1 - top - bottom;
    if (visibleW > 0.001 && visibleH > 0.001) {
      imgStyle.width = `${(1 / visibleW) * 100}%`;
      imgStyle.height = `${(1 / visibleH) * 100}%`;
      imgStyle.marginLeft = `${(-left / visibleW) * 100}%`;
      imgStyle.marginTop = `${(-top / visibleH) * 100}%`;
      imgStyle.objectFit = "fill";
    }
  }

  return (
    <div style={{ ...style, overflow: "hidden" }}>
      <img src={src} alt="" draggable={false} style={imgStyle} />
    </div>
  );
}

// ─── Table Renderer ─────────────────────────────────────────────────────────

function TableRenderer({ node, style }: { node: TableNodeData; style: React.CSSProperties }) {
  return (
    <div style={{ ...style, overflow: "hidden" }}>
      <table style={{ width: "100%", height: "100%", borderCollapse: "collapse", tableLayout: "fixed", fontSize: 11 }}>
        <tbody>
          {node.rows.map((row, ri) => (
            <tr key={ri} style={{ height: row.height }}>
              {row.cells.map((cell, ci) => {
                if (cell.hMerge || cell.vMerge) return null;
                const cellFill = resolveFillColor(cell.properties?.child("solidFill").exists() ? cell.properties!.child("solidFill") : undefined);
                return (
                  <td key={ci}
                    colSpan={cell.gridSpan > 1 ? cell.gridSpan : undefined}
                    rowSpan={cell.rowSpan > 1 ? cell.rowSpan : undefined}
                    style={{ border: "1px solid #d0d0d0", padding: "3px 6px", verticalAlign: "top", background: cellFill ?? undefined }}
                  >
                    {cell.textBody?.paragraphs.map((p, pi) => <ParagraphRenderer key={pi} paragraph={p} />)}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ─── Group Renderer ─────────────────────────────────────────────────────────

function GroupRenderer({ style }: { node?: GroupNodeData; style: React.CSSProperties }) {
  return <div style={style} />;
}
