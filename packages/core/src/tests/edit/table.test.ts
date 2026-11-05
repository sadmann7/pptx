// @vitest-environment jsdom
import { describe, expect, it } from "vitest";

import { applyEdit } from "../../edit/operation";
import type { TableNodeData } from "../../model/nodes/table";
import { buildPresentation, PresentationData } from "../../model/presentation";
import { writePptx } from "../../ooxml/writer";
import { readPptx } from "../../ooxml/zip";
import { buildCustomPptx } from "../fixtures/fixture-extras";

function cell(text: string, attrs = ""): string {
  const run = text ? `<a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r>` : "";
  return `<a:tc${attrs}><a:txBody><a:bodyPr/><a:lstStyle/><a:p>${run}</a:p></a:txBody><a:tcPr/></a:tc>`;
}

function table(id: number, rows: string[][]): string {
  return `<p:graphicFrame>
<p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="740664"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
<a:tblPr/><a:tblGrid><a:gridCol w="914400"/><a:gridCol w="914400"/></a:tblGrid>
${rows.map((row) => `<a:tr h="370332">${row.join("")}</a:tr>`).join("\n")}
</a:tbl></a:graphicData></a:graphic>
</p:graphicFrame>`;
}

async function openEditable(shapes: string[]): Promise<PresentationData> {
  const buffer = await buildCustomPptx({ slides: [shapes.join("\n")] });
  return buildPresentation(await readPptx(buffer, { keepSourcePackage: true }));
}

async function saveAndReopen(pres: PresentationData): Promise<PresentationData> {
  const saved = await writePptx(pres);
  return buildPresentation(await readPptx(saved.slice().buffer));
}

const cellText = (node: TableNodeData, row: number, column: number) =>
  node.rows[row].cells[column].textBody!.paragraphs.map((paragraph) =>
    paragraph.runs.map((run) => run.text).join(""),
  );

describe("addTable", () => {
  it("adds a styled table with empty cells, centered, on top", async () => {
    const pres = await openEditable([""]);
    const slideId = pres.slides[0].id;

    const result = await applyEdit(pres, { type: "addTable", slideId, rows: 3, columns: 4 });

    expect(result.createdNodeIds).toHaveLength(1);
    const node = pres.slides[0].nodes[0] as TableNodeData;
    expect(node.nodeType).toBe("table");
    expect(node.columns).toHaveLength(4);
    expect(node.rows).toHaveLength(3);
    expect(node.size.w).toBeCloseTo((pres.width * 2) / 3, 0);
    expect(node.size.h).toBeCloseTo(120, 0);
    expect(node.position.x).toBeCloseTo((pres.width - node.size.w) / 2, 0);
    expect(node.tableStyleId).toBe("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
    expect(cellText(node, 2, 3)).toEqual([""]);

    const reopened = (await saveAndReopen(pres)).slides[0].nodes[0] as TableNodeData;
    expect(reopened.rows).toHaveLength(3);
    expect(reopened.columns).toHaveLength(4);

    result.undo();
    expect(pres.slides[0].nodes).toHaveLength(0);
  });

  it("takes size and position, and rejects bad dimensions", async () => {
    const pres = await openEditable([""]);
    const slideId = pres.slides[0].id;
    await applyEdit(pres, {
      type: "addTable",
      slideId,
      rows: 2,
      columns: 2,
      size: { w: 200, h: 80 },
      position: { x: 10, y: 20 },
    });

    const node = pres.slides[0].nodes[0] as TableNodeData;
    expect(node.size.w).toBeCloseTo(200, 0);
    expect(node.size.h).toBeCloseTo(80, 0);
    expect(node.position).toEqual({ x: 10, y: 20 });
    await expect(
      applyEdit(pres, { type: "addTable", slideId, rows: 0, columns: 2 }),
    ).rejects.toThrow("invalid table size");
    await expect(
      applyEdit(pres, { type: "addTable", slideId, rows: 2, columns: 1.5 }),
    ).rejects.toThrow("invalid table size");
  });
});

describe("table cell text", () => {
  it("replaces one cell's text, saves, and undoes", async () => {
    const pres = await openEditable([
      table(2, [
        [cell("A"), cell("B")],
        [cell("C"), cell("")],
      ]),
    ]);
    const slideId = pres.slides[0].id;

    const result = await applyEdit(pres, {
      type: "setTextBody",
      slideId,
      nodeId: "2",
      cell: { row: 1, column: 1 },
      paragraphs: [{ sourceParagraphIndex: 0, runs: [{ text: "D" }] }],
    });

    const node = pres.slides[0].nodes[0] as TableNodeData;
    expect(cellText(node, 1, 1)).toEqual(["D"]);
    expect(cellText(node, 0, 0)).toEqual(["A"]);

    const reopened = (await saveAndReopen(pres)).slides[0].nodes[0] as TableNodeData;
    expect(cellText(reopened, 1, 1)).toEqual(["D"]);

    result.undo();
    expect(cellText(node, 1, 1)).toEqual([""]);
  });

  it("formats and aligns a cell's text", async () => {
    const pres = await openEditable([table(2, [[cell("Hello"), cell("B")]])]);
    const slideId = pres.slides[0].id;
    const address = { row: 0, column: 0 };

    await applyEdit(pres, {
      type: "batch",
      operations: [
        { type: "setTextStyle", slideId, nodeId: "2", cell: address, style: { bold: true } },
        {
          type: "setParagraphStyle",
          slideId,
          nodeId: "2",
          cell: address,
          style: { align: "center" },
        },
      ],
    });

    const node = pres.slides[0].nodes[0] as TableNodeData;
    const paragraph = node.rows[0].cells[0].textBody!.paragraphs[0];
    expect(paragraph.runs[0].properties?.attr("b")).toBe("1");
    expect(paragraph.properties?.attr("algn")).toBe("ctr");
    expect(
      node.rows[0].cells[1].textBody!.paragraphs[0].runs[0].properties?.attr("b"),
    ).toBeUndefined();
  });

  it("rejects merged continuation cells and cells on shapes", async () => {
    const pres = await openEditable([
      table(2, [[cell("A", ' gridSpan="2"'), cell("", ' hMerge="1"')]]),
    ]);
    const slideId = pres.slides[0].id;

    await expect(
      applyEdit(pres, {
        type: "setTextBody",
        slideId,
        nodeId: "2",
        cell: { row: 0, column: 1 },
        paragraphs: [{ sourceParagraphIndex: 0, runs: [{ text: "X" }] }],
      }),
    ).rejects.toThrow("no editable cell");
    await expect(
      applyEdit(pres, {
        type: "setTextBody",
        slideId,
        nodeId: "2",
        paragraphs: [{ sourceParagraphIndex: 0, runs: [{ text: "X" }] }],
      }),
    ).rejects.toThrow("no editable text body");
  });
});
