/**
 * Shared helpers for the text-rendering test suites.
 *
 * All helpers run the REAL pipeline (parseZip → buildPresentation → renderSlide)
 * against in-memory fixtures produced by `buildPptxWithShapes` — no mocks.
 */
import { buildPresentation } from "../../model/presentation";
import { parseZip } from "../../ooxml/zip-parser";
import { renderSlide, type SlideRendererOptions } from "../../renderer/slide-renderer";
import { buildPptxWithShapes } from "./minimal-pptx";

/** Render arbitrary spTree shapes XML through the full pipeline. */
export async function renderShapes(
  shapesXml: string,
  options?: SlideRendererOptions,
): Promise<HTMLElement> {
  const buffer = await buildPptxWithShapes(shapesXml);
  const files = await parseZip(buffer);
  const presentation = buildPresentation(files);
  return renderSlide(presentation, presentation.slides[0], options).element;
}

/** Default bodyPr matching the existing text-renderer.test.ts fixture (noAutofit). */
export const DEFAULT_BODY_PR = `<a:bodyPr anchorCtr="0" anchor="t" bIns="91425" lIns="91425" spcFirstLastPara="1" rIns="91425" wrap="square" tIns="91425"><a:noAutofit/></a:bodyPr>`;

interface RenderTextBoxOptions {
  /** Full `<a:bodyPr …>…</a:bodyPr>` element; defaults to the noAutofit fixture bodyPr. */
  bodyPrXml?: string;
  /** Full `<a:lstStyle …>…</a:lstStyle>` element; defaults to an empty lstStyle. */
  lstStyleXml?: string;
  /** Contents of `<p:nvPr>` (e.g. `<p:ph type="title"/>`). */
  nvPrXml?: string;
  renderOptions?: SlideRendererOptions;
}

/** Renders a single text box with the given txBody paragraphs through the full pipeline. */
export async function renderTextBox(
  paragraphsXml: string,
  options?: RenderTextBoxOptions,
): Promise<HTMLElement> {
  return renderShapes(
    `<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr>${options?.nvPrXml ?? ""}</p:nvPr></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="251751" y="3062425"/><a:ext cx="7788000" cy="485100"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
</p:spPr>
<p:txBody>
${options?.bodyPrXml ?? DEFAULT_BODY_PR}
${options?.lstStyleXml ?? "<a:lstStyle/>"}
${paragraphsXml}
</p:txBody>
</p:sp>`,
    options?.renderOptions,
  );
}

/**
 * The text container is the flex-column div that renderTextBody fills with
 * paragraph divs (created by the shape renderer for the single fixture shape).
 */
export function textContainerOf(element: HTMLElement): HTMLElement {
  const isContainer = (div: HTMLElement) =>
    div.style.display === "flex" && div.style.flexDirection === "column";
  if (isContainer(element)) return element;
  const container = [...element.querySelectorAll("div")].find((div) =>
    isContainer(div as HTMLElement),
  );
  if (!container) throw new Error("text container not found");
  return container as HTMLElement;
}

/** Paragraph divs are the direct children of the text container. */
export function paragraphsOf(element: HTMLElement): HTMLElement[] {
  return [...textContainerOf(element).children] as HTMLElement[];
}

/** Non-empty spans of a paragraph div (runs and bullets). */
export function spansOf(paragraph: HTMLElement): HTMLElement[] {
  return [...paragraph.querySelectorAll("span")] as HTMLElement[];
}

/**
 * Normalize a CSS color ("#RRGGBB", "#rgb" or "rgb()/rgba()") to a comparable
 * "r,g,b,a" string, so assertions survive engine-specific serialization.
 */
export function normalizeColor(color: string | undefined): string | undefined {
  if (!color) return undefined;
  const trimmed = color.trim().toLowerCase();
  const hexMatch = trimmed.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/);
  if (hexMatch) {
    let hex = hexMatch[1];
    if (hex.length === 3) {
      hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
    }
    const num = parseInt(hex, 16);
    return `${(num >> 16) & 0xff},${(num >> 8) & 0xff},${num & 0xff},1`;
  }
  const rgbMatch = trimmed.match(/^rgba?\(([^)]+)\)$/);
  if (rgbMatch) {
    const parts = rgbMatch[1].split(",").map((part) => Number(part.trim()));
    const [r, g, b] = parts;
    const a = parts.length > 3 ? parts[3] : 1;
    return `${r},${g},${b},${a}`;
  }
  return trimmed;
}
