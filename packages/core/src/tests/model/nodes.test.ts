import { beforeAll, describe, expect, it } from "vitest";

import type { GroupNodeData } from "../../model/nodes/group";
import type { PictureNodeData } from "../../model/nodes/picture";
import type { PresentationData } from "../../model/presentation";
import { buildPresentation, materializeAllSlides } from "../../model/presentation";
import { readPptx } from "../../ooxml/zip";
import { renderSlide } from "../../renderer/slide";
import { buildPptxWithShapes } from "../fixtures/minimal-pptx";

const GROUPED_SHAPES = `<p:grpSp>
<p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr>
<a:xfrm>
<a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/>
<a:chOff x="0" y="0"/><a:chExt cx="914400" cy="457200"/>
</a:xfrm>
</p:grpSpPr>
<p:sp>
<p:nvSpPr><p:cNvPr id="11" name="Inner"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill>
</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>
</p:grpSp>`;

const PICTURE = `<p:pic>
<p:nvPicPr>
<p:cNvPr id="20" name="Photo"/>
<p:cNvPicPr/>
<p:nvPr/>
</p:nvPicPr>
<p:blipFill>
<a:blip r:embed="rId9"/>
<a:srcRect t="10000" b="20000" l="5000" r="0"/>
<a:stretch><a:fillRect/></a:stretch>
</p:blipFill>
<p:spPr>
<a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
</p:spPr>
</p:pic>`;

const VIDEO = `<p:pic>
<p:nvPicPr>
<p:cNvPr id="21" name="Clip"/>
<p:cNvPicPr/>
<p:nvPr><a:videoFile r:link="rId10"/></p:nvPr>
</p:nvPicPr>
<p:blipFill><a:blip r:embed="rId11"/></p:blipFill>
<p:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
</p:spPr>
</p:pic>`;

let presentation: PresentationData;

beforeAll(async () => {
  const buffer = await buildPptxWithShapes(GROUPED_SHAPES + PICTURE + VIDEO);
  presentation = buildPresentation(await readPptx(buffer));
  materializeAllSlides(presentation);
});

describe("group nodes", () => {
  it("parses group child coordinate space and children", () => {
    const group = presentation.slides[0].nodes.find(
      (n): n is GroupNodeData => n.nodeType === "group",
    );
    expect(group).toBeDefined();
    expect(group!.childOffset).toEqual({ x: 0, y: 0 });
    // chExt 914400x457200 EMU = 96x48 px
    expect(group!.childExtent).toEqual({ w: 96, h: 48 });
    expect(group!.children).toHaveLength(1);
    expect(group!.children[0].localName).toBe("sp");
  });

  it("renders grouped children scaled into the group box", () => {
    const element = renderSlide(presentation, presentation.slides[0]).element;
    // The inner orange rect must appear in the rendered slide.
    const paths = [...element.querySelectorAll("svg path")];
    const orange = paths.find((p) => p.getAttribute("fill")?.toLowerCase() === "#ed7d31");
    expect(orange).toBeDefined();
  });
});

describe("picture nodes", () => {
  it("parses blip embed, crop, and clip geometry", () => {
    const pic = presentation.slides[0].nodes.find(
      (n): n is PictureNodeData => n.nodeType === "picture" && !n.isVideo,
    );
    expect(pic).toBeDefined();
    expect(pic!.blipEmbed).toBe("rId9");
    expect(pic!.presetGeometry).toBe("ellipse");
    expect(pic!.crop).toEqual({ top: 0.1, bottom: 0.2, left: 0.05, right: 0 });
  });

  it("detects video placeholders with their media relationship", () => {
    const video = presentation.slides[0].nodes.find(
      (n): n is PictureNodeData => n.nodeType === "picture" && n.isVideo === true,
    );
    expect(video).toBeDefined();
    expect(video!.mediaRId).toBe("rId10");
  });
});
