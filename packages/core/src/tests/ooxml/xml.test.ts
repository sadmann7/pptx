import { describe, expect, it } from "vitest";

import { parseXml, SafeXmlNode } from "../../ooxml/xml";

const SLIDE_XML = `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:spPr>
          <a:xfrm rot="60000"><a:off x="10" y="20"/><a:ext cx="100" cy="200"/></a:xfrm>
          <a:blipFill><a:blip r:embed="rId7"/></a:blipFill>
        </p:spPr>
        <p:txBody><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:pic/>
      <p:sp/>
    </p:spTree>
  </p:cSld>
</p:sld>`;

describe("parseXml", () => {
  it("parses a document and exposes the root element", () => {
    const root = parseXml(SLIDE_XML);
    expect(root.exists()).toBe(true);
    expect(root.localName).toBe("sld");
  });

  it("returns a non-existent node for malformed XML", () => {
    const root = parseXml("<oops <<<");
    expect(root.exists()).toBe(false);
  });

  it("strips a leading UTF-8 BOM before parsing", () => {
    const root = parseXml(`\uFEFF${SLIDE_XML}`);
    expect(root.exists()).toBe(true);
    expect(root.localName).toBe("sld");
  });
});

describe("SafeXmlNode", () => {
  const root = parseXml(SLIDE_XML);
  const spTree = root.child("cSld").child("spTree");

  it("navigates children by localName ignoring namespace prefixes", () => {
    expect(spTree.exists()).toBe(true);
    expect(spTree.child("sp").exists()).toBe(true);
    // a:xfrm found under p:spPr even though prefixes differ
    expect(spTree.child("sp").child("spPr").child("xfrm").exists()).toBe(true);
  });

  it("never crashes when chaining through missing elements", () => {
    const missing = root.child("nope").child("deeper").child("deepest");
    expect(missing.exists()).toBe(false);
    expect(missing.attr("x")).toBeUndefined();
    expect(missing.numAttr("x")).toBeUndefined();
    expect(missing.text()).toBe("");
    expect(missing.children()).toEqual([]);
    expect(missing.localName).toBe("");
  });

  it("reads attributes and numeric attributes", () => {
    const xfrm = spTree.child("sp").child("spPr").child("xfrm");
    expect(xfrm.attr("rot")).toBe("60000");
    expect(xfrm.numAttr("rot")).toBe(60000);
    expect(xfrm.child("off").numAttr("x")).toBe(10);
    expect(xfrm.child("off").numAttr("missing")).toBeUndefined();
  });

  it("returns undefined for non-numeric attribute values via numAttr", () => {
    const node = parseXml(`<n val="abc"/>`);
    expect(node.attr("val")).toBe("abc");
    expect(node.numAttr("val")).toBeUndefined();
  });

  it("resolves namespace-prefixed attributes like r:embed", () => {
    const blip = spTree.child("sp").child("spPr").child("blipFill").child("blip");
    expect(blip.attr("r:embed")).toBe("rId7");
  });

  it("filters children by localName and counts all children", () => {
    expect(spTree.children("sp")).toHaveLength(2);
    expect(spTree.children("pic")).toHaveLength(1);
    expect(spTree.children()).toHaveLength(3);
    expect(spTree.allChildren()).toHaveLength(3);
  });

  it("reads text content through nesting", () => {
    const txBody = spTree.child("sp").child("txBody");
    expect(txBody.text()).toBe("Hello");
  });

  it("wraps null elements safely", () => {
    const node = new SafeXmlNode(null);
    expect(node.exists()).toBe(false);
    expect(node.element).toBeNull();
  });
});
