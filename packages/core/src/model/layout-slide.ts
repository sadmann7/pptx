import { A_NS, P_NS, R_NS, serializePartText } from "../edit/xml-mutation";
import { parseXml, SafeXmlNode } from "../ooxml/xml";
import type { PresentationData } from "./presentation";
import { parseSlide, SlideData } from "./slide";

// PowerPoint leaves header and footer placeholders off a new slide.
export const SKIPPED_PLACEHOLDER_TYPES = new Set([
  "dt",
  "ftr",
  "sldNum",
  "hdr",
]);

const EMPTY_SLIDE_XML =
  `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
  "<p:cSld><p:spTree>" +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
  "</p:spTree></p:cSld>" +
  "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>" +
  "</p:sld>";

// The copy carries only the `p:ph` reference, so position, size, and text
// styles keep inheriting from the layout.
export function createPlaceholderFromLayout(
  doc: Document,
  layoutSp: SafeXmlNode,
  id: number,
): Element | null {
  const nvSpPr = layoutSp.child("nvSpPr");
  const phElement = nvSpPr.child("nvPr").child("ph").element;
  if (!phElement) return null;

  const el = (
    qualifiedName: string,
    attrs: Record<string, string> = {},
    ...children: Node[]
  ) => {
    const element = doc.createElementNS(
      qualifiedName.startsWith("a:") ? A_NS : P_NS,
      qualifiedName,
    );
    for (const [name, value] of Object.entries(attrs))
      element.setAttribute(name, value);
    for (const child of children) element.appendChild(child);
    return element;
  };

  const phCopy = doc.importNode(phElement, true) as Element;
  phCopy.removeAttribute("hasCustomPrompt");
  const name = nvSpPr.child("cNvPr").attr("name") ?? `Placeholder ${id - 1}`;

  return el(
    "p:sp",
    {},
    el(
      "p:nvSpPr",
      {},
      el("p:cNvPr", { id: String(id), name }),
      el("p:cNvSpPr", {}, el("a:spLocks", { noGrp: "1" })),
      el("p:nvPr", {}, phCopy),
    ),
    el("p:spPr"),
    el(
      "p:txBody",
      {},
      el("a:bodyPr"),
      el("a:lstStyle"),
      el("a:p", {}, el("a:endParaRPr", { lang: "en-US" })),
    ),
  );
}

export function buildSlideFromLayout(layoutSpTree: SafeXmlNode): string {
  const root = parseXml(EMPTY_SLIDE_XML);
  const rootElement = root.element!;
  const spTree = root.child("cSld").child("spTree").element!;

  let nextId = 2;
  for (const sp of layoutSpTree.children("sp")) {
    const type = sp.child("nvSpPr").child("nvPr").child("ph").attr("type");
    if (SKIPPED_PLACEHOLDER_TYPES.has(type ?? "")) continue;
    const placeholder = createPlaceholderFromLayout(
      rootElement.ownerDocument,
      sp,
      nextId,
    );
    if (!placeholder) continue;
    spTree.appendChild(placeholder);
    nextId++;
  }

  return serializePartText(rootElement, "");
}

// Detached from the deck: it has no part, so nothing about it is saved.
export function createLayoutSlide(
  pres: PresentationData,
  layoutId: string,
): SlideData {
  const layout = pres.layouts.get(layoutId);
  if (!layout) throw new Error(`unknown slide layout "${layoutId}".`);
  const slide = parseSlide(
    parseXml(buildSlideFromLayout(layout.spTree)),
    -1,
    new Map(),
    "",
    pres.diagramDrawings,
  );
  slide.layoutIndex = layoutId;
  return slide;
}
