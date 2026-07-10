/**
 * Picture rendering (image-renderer) with a real embedded PNG resolved
 * through slide rels and zip media parts.
 *
 * happy-dom performs no image decoding, so assertions target the synchronous
 * DOM output: <img> elements, blob URLs, crop scaling, and SVG clip structure.
 */
import { describe, expect, it } from "vitest";

import { buildPresentation } from "../../model/presentation";
import { parseZip } from "../../ooxml/zip-parser";
import { renderSlide } from "../../renderer/slide-renderer";
import { buildRichPptx, tinyPngBytes } from "../helpers/rich-pptx";

const IMAGE_REL = `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>`;

async function renderPicture(picXml: string): Promise<HTMLElement> {
  const buffer = await buildRichPptx({
    shapesXml: picXml,
    extraSlideRelsXml: IMAGE_REL,
    binaryParts: { "ppt/media/image1.png": tinyPngBytes() },
  });
  const files = await parseZip(buffer);
  const presentation = buildPresentation(files);
  return renderSlide(presentation, presentation.slides[0]).element;
}

function pic(blipFillInner: string, spPrInner: string, xfrmAttrs = ""): string {
  return `<p:pic>
<p:nvPicPr><p:cNvPr id="20" name="Photo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
<p:blipFill>
${blipFillInner}
</p:blipFill>
<p:spPr>
<a:xfrm${xfrmAttrs}><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
${spPrInner}
</p:spPr>
</p:pic>`;
}

const SIMPLE_BLIP = `<a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch>`;
const RECT_GEOM = `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`;

describe("basic picture rendering", () => {
  it("renders an embedded PNG as an <img> with a blob URL", async () => {
    const element = await renderPicture(pic(SIMPLE_BLIP, RECT_GEOM));

    const img = element.querySelector("img")!;
    expect(img).not.toBeNull();
    expect(img.getAttribute("src")).toMatch(/^blob:/);
    expect(img.style.objectFit).toBe("fill");

    const wrapper = img.closest("div") as HTMLElement;
    expect(wrapper.style.left).toBe("96px");
    expect(wrapper.style.top).toBe("96px");
    expect(wrapper.style.width).toBe("96px");
    expect(wrapper.style.height).toBe("96px");
    expect(wrapper.style.overflow).toBe("hidden");
  });

  it("reuses the cached blob URL for the same media across renders", async () => {
    const buffer = await buildRichPptx({
      shapesXml: pic(SIMPLE_BLIP, RECT_GEOM),
      extraSlideRelsXml: IMAGE_REL,
      binaryParts: { "ppt/media/image1.png": tinyPngBytes() },
    });
    const files = await parseZip(buffer);
    const presentation = buildPresentation(files);
    const cache = new Map<string, string>();
    const first = renderSlide(presentation, presentation.slides[0], {
      mediaUrlCache: cache,
    }).element;
    const second = renderSlide(presentation, presentation.slides[0], {
      mediaUrlCache: cache,
    }).element;
    expect(first.querySelector("img")!.getAttribute("src")).toBe(
      second.querySelector("img")!.getAttribute("src"),
    );
  });

  it("renders a placeholder when the blip relationship is missing", async () => {
    const element = await renderPicture(
      pic(`<a:blip r:embed="rId99"/><a:stretch><a:fillRect/></a:stretch>`, RECT_GEOM),
    );
    expect(element.querySelector("img")).toBeNull();
    expect(element.textContent).toContain("Missing image reference");
  });
});

describe("picture crop (a:srcRect)", () => {
  it("scales and offsets the image so the visible crop fills the shape box", async () => {
    // 10% cropped from left and right: visible width 80% -> scale 1.25.
    const element = await renderPicture(
      pic(
        `<a:blip r:embed="rId9"/><a:srcRect l="10000" r="10000"/>${"<a:stretch><a:fillRect/></a:stretch>"}`,
        RECT_GEOM,
      ),
    );
    const img = element.querySelector("img")!;
    // Wrapper is 96px wide: img 1.25*96 = 120px, offset -0.1*1.25*96 = -12px.
    expect(parseFloat(img.style.width)).toBeCloseTo(120, 3);
    expect(parseFloat(img.style.height)).toBeCloseTo(96, 3);
    expect(parseFloat(img.style.marginLeft)).toBeCloseTo(-12, 3);
    expect(parseFloat(img.style.marginTop)).toBeCloseTo(0, 3);
  });
});

describe("picture geometry clipping", () => {
  it("renders non-rect preset geometry as an SVG-clipped image", async () => {
    const element = await renderPicture(
      pic(SIMPLE_BLIP, `<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>`),
    );

    // No plain <img>: the picture becomes an SVG <image> clipped by the preset path.
    expect(element.querySelector("img")).toBeNull();
    const svg = element.querySelector("svg")!;
    expect(svg).not.toBeNull();
    const clipPath = svg.querySelector("defs clipPath")!;
    expect(clipPath).not.toBeNull();
    expect(clipPath.querySelector("path")!.getAttribute("d")).toBeTruthy();

    const group = svg.querySelector("g")!;
    expect(group.getAttribute("clip-path")).toMatch(/^url\(#picture-clip-/);

    const image = svg.querySelector("image")!;
    expect(image.getAttribute("href")).toMatch(/^blob:/);
    expect(Number(image.getAttribute("width"))).toBeCloseTo(96, 5);
    expect(Number(image.getAttribute("height"))).toBeCloseTo(96, 5);
    expect(image.getAttribute("preserveAspectRatio")).toBe("none");
  });

  it("applies crop scaling inside the SVG clip structure", async () => {
    const element = await renderPicture(
      pic(
        `<a:blip r:embed="rId9"/><a:srcRect t="10000" b="20000" l="5000" r="0"/><a:stretch><a:fillRect/></a:stretch>`,
        `<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>`,
      ),
    );
    const image = element.querySelector("svg image")!;
    // visibleW = 0.95 -> width 96/0.95; visibleH = 0.7 -> height 96/0.7.
    expect(Number(image.getAttribute("width"))).toBeCloseTo(96 / 0.95, 3);
    expect(Number(image.getAttribute("height"))).toBeCloseTo(96 / 0.7, 3);
    expect(Number(image.getAttribute("x"))).toBeCloseTo(-0.05 * (1 / 0.95) * 96, 3);
    expect(Number(image.getAttribute("y"))).toBeCloseTo(-0.1 * (1 / 0.7) * 96, 3);
  });
});

describe("picture transforms and effects", () => {
  it("applies flipH as scaleX(-1) on rectangular pictures", async () => {
    const element = await renderPicture(pic(SIMPLE_BLIP, RECT_GEOM, ` flipH="1"`));
    const wrapper = element.querySelector("img")!.closest("div") as HTMLElement;
    expect(wrapper.style.transform).toBe("scaleX(-1)");
  });

  it("flips clipped pictures via SVG transforms instead of CSS scale", async () => {
    const element = await renderPicture(
      pic(SIMPLE_BLIP, `<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>`, ` flipH="1"`),
    );
    const wrapper = element.querySelector("svg")!.closest("div") as HTMLElement;
    expect(wrapper.style.transform).not.toContain("scaleX");
    const clipPathPath = element.querySelector("svg defs clipPath path")!;
    expect(clipPathPath.getAttribute("transform")).toBe("translate(96 0) scale(-1 1)");
    const image = element.querySelector("svg image")!;
    expect(image.getAttribute("transform")).toBe("translate(96 0) scale(-1 1)");
  });

  it("applies blip alphaModFix as wrapper opacity", async () => {
    const element = await renderPicture(
      pic(
        `<a:blip r:embed="rId9"><a:alphaModFix amt="50000"/></a:blip><a:stretch><a:fillRect/></a:stretch>`,
        RECT_GEOM,
      ),
    );
    const wrapper = element.querySelector("img")!.closest("div") as HTMLElement;
    expect(wrapper.style.opacity).toBe("0.5");
  });

  it("applies an spPr outline as a wrapper border", async () => {
    const element = await renderPicture(
      pic(
        SIMPLE_BLIP,
        `${RECT_GEOM}<a:ln w="25400"><a:solidFill><a:srgbClr val="1F1F1F"/></a:solidFill></a:ln>`,
      ),
    );
    const wrapper = element.querySelector("img")!.closest("div") as HTMLElement;
    expect(wrapper.style.border).toMatch(/solid/);
    expect(wrapper.style.border).toMatch(/#1f1f1f/i);
    expect(wrapper.style.boxSizing).toBe("border-box");
  });
});
