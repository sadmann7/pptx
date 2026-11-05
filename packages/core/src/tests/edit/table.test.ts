// @vitest-environment jsdom
import { describe, expect, it } from "vitest";

import { applyEdit, getTableCellRange } from "../../edit/operation";
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
<a:tblPr/><a:tblGrid>${rows[0].map(() => '<a:gridCol w="914400"/>').join("")}</a:tblGrid>
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

const texts = (node: TableNodeData) =>
  node.rows.map((row) =>
    row.cells.map((_, column) => cellText(node, node.rows.indexOf(row), column).join("\n")),
  );

const tableOf = (pres: PresentationData) => pres.slides[0].nodes[0] as TableNodeData;

describe("table rows and columns", () => {
  it("inserts an empty row, saves, and undoes", async () => {
    const pres = await openEditable([
      table(2, [
        [cell("A"), cell("B")],
        [cell("C"), cell("D")],
      ]),
    ]);
    const slideId = pres.slides[0].id;

    const result = await applyEdit(pres, {
      type: "insertTableRow",
      slideId,
      nodeId: "2",
      index: 1,
    });

    expect(texts(tableOf(pres))).toEqual([
      ["A", "B"],
      ["", ""],
      ["C", "D"],
    ]);
    expect(tableOf(pres).size.h).toBeCloseTo((370332 * 3) / 9525, 0);
    const paragraph = tableOf(pres).rows[1].cells[0].textBody!.paragraphs[0];
    expect(paragraph.endParaRPr?.attr("lang")).toBe("en-US");

    const reopened = (await saveAndReopen(pres)).slides[0].nodes[0] as TableNodeData;
    expect(texts(reopened)).toEqual([
      ["A", "B"],
      ["", ""],
      ["C", "D"],
    ]);

    result.undo();
    expect(texts(tableOf(pres))).toEqual([
      ["A", "B"],
      ["C", "D"],
    ]);
  });

  it("inserts and deletes columns", async () => {
    const pres = await openEditable([
      table(2, [
        [cell("A"), cell("B")],
        [cell("C"), cell("D")],
      ]),
    ]);
    const slideId = pres.slides[0].id;

    await applyEdit(pres, { type: "insertTableColumn", slideId, nodeId: "2", index: 2 });
    expect(texts(tableOf(pres))).toEqual([
      ["A", "B", ""],
      ["C", "D", ""],
    ]);
    expect(tableOf(pres).columns).toHaveLength(3);

    await applyEdit(pres, { type: "deleteTableColumn", slideId, nodeId: "2", index: 0 });
    expect(texts(tableOf(pres))).toEqual([
      ["B", ""],
      ["D", ""],
    ]);
    expect(tableOf(pres).columns).toHaveLength(2);

    await applyEdit(pres, { type: "deleteTableRow", slideId, nodeId: "2", index: 0 });
    expect(texts(tableOf(pres))).toEqual([["D", ""]]);
  });

  it("grows and shrinks merges that span the change", async () => {
    const pres = await openEditable([
      table(2, [
        [cell("Merged", ' rowSpan="2"'), cell("B")],
        [cell("", ' vMerge="1"'), cell("D")],
      ]),
    ]);
    const slideId = pres.slides[0].id;

    await applyEdit(pres, { type: "insertTableRow", slideId, nodeId: "2", index: 1 });
    let node = tableOf(pres);
    expect(node.rows[0].cells[0].rowSpan).toBe(3);
    expect(node.rows[1].cells[0].vMerge).toBe(true);
    expect(node.rows[1].cells[1].vMerge).toBe(false);

    await applyEdit(pres, { type: "deleteTableRow", slideId, nodeId: "2", index: 0 });
    node = tableOf(pres);
    expect(node.rows[0].cells[0].rowSpan).toBe(2);
    expect(node.rows[0].cells[0].vMerge).toBe(false);
    expect(cellText(node, 0, 0)).toEqual(["Merged"]);
  });

  it("grows a horizontal merge when a column goes inside it", async () => {
    const pres = await openEditable([
      table(2, [
        [cell("Wide", ' gridSpan="2"'), cell("", ' hMerge="1"')],
        [cell("C"), cell("D")],
      ]),
    ]);
    const slideId = pres.slides[0].id;

    await applyEdit(pres, { type: "insertTableColumn", slideId, nodeId: "2", index: 1 });
    const node = tableOf(pres);
    expect(node.rows[0].cells[0].gridSpan).toBe(3);
    expect(node.rows[0].cells[1].hMerge).toBe(true);
    expect(texts(node)[1]).toEqual(["C", "", "D"]);
  });

  it("rejects bad indexes and removing the last row or column", async () => {
    const pres = await openEditable([table(2, [[cell("A")]])]);
    const slideId = pres.slides[0].id;

    await expect(
      applyEdit(pres, { type: "insertTableRow", slideId, nodeId: "2", index: 5 }),
    ).rejects.toThrow("out of bounds");
    await expect(
      applyEdit(pres, { type: "deleteTableRow", slideId, nodeId: "2", index: 0 }),
    ).rejects.toThrow("last row");
    await expect(
      applyEdit(pres, { type: "deleteTableColumn", slideId, nodeId: "2", index: 0 }),
    ).rejects.toThrow("last column");
  });
});

describe("setTableGrid", () => {
  it("sets column widths and row heights, saves, and undoes", async () => {
    const pres = await openEditable([
      table(2, [
        [cell("A"), cell("B")],
        [cell("C"), cell("D")],
      ]),
    ]);
    const slideId = pres.slides[0].id;
    const before = { columns: [...tableOf(pres).columns], size: { ...tableOf(pres).size } };

    const result = await applyEdit(pres, {
      type: "setTableGrid",
      slideId,
      nodeId: "2",
      columnWidths: [60, 132],
      rowHeights: [50, 39],
    });

    expect(tableOf(pres).columns.map(Math.round)).toEqual([60, 132]);
    expect(tableOf(pres).rows.map((row) => Math.round(row.height))).toEqual([50, 39]);
    expect(Math.round(tableOf(pres).size.h)).toBe(89);
    const reopened = (await saveAndReopen(pres)).slides[0].nodes[0] as TableNodeData;
    expect(reopened.columns.map(Math.round)).toEqual([60, 132]);

    result.undo();
    expect(tableOf(pres).columns).toEqual(before.columns);
    expect(tableOf(pres).size).toEqual(before.size);
  });

  it("rejects sizes that don't match the grid", async () => {
    const pres = await openEditable([table(2, [[cell("A"), cell("B")]])]);
    await expect(
      applyEdit(pres, {
        type: "setTableGrid",
        slideId: pres.slides[0].id,
        nodeId: "2",
        columnWidths: [100],
      }),
    ).rejects.toThrow("don't match");
  });
});

describe("merging table cells", () => {
  const grid = () =>
    table(2, [
      [cell("A"), cell("B"), cell("C")],
      [cell("D"), cell(""), cell("F")],
      [cell("G"), cell("H"), cell("I")],
    ]);

  it("merges a rectangle, keeping its text in the top-left cell, and splits it back", async () => {
    const pres = await openEditable([grid()]);
    const slideId = pres.slides[0].id;

    const merge = await applyEdit(pres, {
      type: "mergeTableCells",
      slideId,
      nodeId: "2",
      from: { row: 1, column: 1 },
      to: { row: 0, column: 0 },
    });

    let node = tableOf(pres);
    expect(node.rows[0].cells[0].gridSpan).toBe(2);
    expect(node.rows[0].cells[0].rowSpan).toBe(2);
    expect(node.rows[0].cells[1].hMerge).toBe(true);
    expect(node.rows[1].cells[0].vMerge).toBe(true);
    expect(node.rows[1].cells[1].hMerge && node.rows[1].cells[1].vMerge).toBe(true);
    expect(cellText(node, 0, 0)).toEqual(["A", "B", "D"]);
    const reopened = (await saveAndReopen(pres)).slides[0].nodes[0] as TableNodeData;
    expect(reopened.rows[0].cells[0].gridSpan).toBe(2);

    const split = await applyEdit(pres, {
      type: "splitTableCell",
      slideId,
      nodeId: "2",
      cell: { row: 1, column: 1 },
    });
    node = tableOf(pres);
    expect(
      node.rows.flatMap((row) => row.cells.map((c) => [c.gridSpan, c.rowSpan, c.hMerge, c.vMerge])),
    ).toEqual(Array.from({ length: 9 }, () => [1, 1, false, false]));
    expect(cellText(node, 0, 0)).toEqual(["A", "B", "D"]);
    expect(cellText(node, 0, 1)).toEqual([""]);

    split.undo();
    merge.undo();
    expect(texts(tableOf(pres))).toEqual([
      ["A", "B", "C"],
      ["D", "", "F"],
      ["G", "H", "I"],
    ]);
  });

  it("grows a range to cover a merge it cuts through", async () => {
    const pres = await openEditable([grid()]);
    const slideId = pres.slides[0].id;
    await applyEdit(pres, {
      type: "mergeTableCells",
      slideId,
      nodeId: "2",
      from: { row: 0, column: 1 },
      to: { row: 1, column: 1 },
    });

    expect(getTableCellRange(tableOf(pres), { row: 1, column: 0 }, { row: 1, column: 1 })).toEqual({
      start: { row: 0, column: 0 },
      end: { row: 1, column: 1 },
    });
  });

  it("rejects merging one cell and splitting a plain cell", async () => {
    const pres = await openEditable([grid()]);
    const slideId = pres.slides[0].id;
    await expect(
      applyEdit(pres, {
        type: "mergeTableCells",
        slideId,
        nodeId: "2",
        from: { row: 0, column: 0 },
        to: { row: 0, column: 0 },
      }),
    ).rejects.toThrow("more than one cell");
    await expect(
      applyEdit(pres, {
        type: "splitTableCell",
        slideId,
        nodeId: "2",
        cell: { row: 0, column: 0 },
      }),
    ).rejects.toThrow("no merged cell");
  });
});

describe("table cell style", () => {
  const twoByTwo = () =>
    table(2, [
      [cell("A"), cell("B")],
      [cell("C"), cell("D")],
    ]);

  it("fills, aligns, and borders a range, saves, and undoes", async () => {
    const pres = await openEditable([twoByTwo()]);
    const slideId = pres.slides[0].id;

    const result = await applyEdit(pres, {
      type: "setTableCellStyle",
      slideId,
      nodeId: "2",
      from: { row: 0, column: 0 },
      to: { row: 0, column: 0 },
      style: { fill: "#ff0000", verticalAlign: "middle", borders: { color: "00ff00", width: 2 } },
    });

    const tcPr = (row: number, column: number) => tableOf(pres).rows[row].cells[column].properties;
    expect(tcPr(0, 0)?.child("solidFill").child("srgbClr").attr("val")).toBe("FF0000");
    expect(tcPr(0, 0)?.attr("anchor")).toBe("ctr");
    expect(tcPr(0, 0)?.child("lnR").attr("w")).toBe("25400");
    expect(
      tcPr(0, 0)
        ?.children()
        .map((child) => child.localName),
    ).toEqual(["lnL", "lnR", "lnT", "lnB", "solidFill"]);
    expect(tcPr(0, 1)?.child("lnL").child("solidFill").exists()).toBe(true);
    expect(tcPr(1, 0)?.child("lnT").child("solidFill").exists()).toBe(true);
    expect(tcPr(1, 1)?.child("lnL").exists()).toBe(false);

    const reopened = (await saveAndReopen(pres)).slides[0].nodes[0] as TableNodeData;
    expect(reopened.rows[0].cells[0].properties?.attr("anchor")).toBe("ctr");

    result.undo();
    expect(tcPr(0, 0)?.child("solidFill").exists()).toBe(false);
  });

  it("formats the whole table without a range, and clears fills and borders", async () => {
    const pres = await openEditable([twoByTwo()]);
    const slideId = pres.slides[0].id;
    await applyEdit(pres, {
      type: "setTableCellStyle",
      slideId,
      nodeId: "2",
      style: { fill: null, borders: null },
    });
    for (const row of tableOf(pres).rows) {
      for (const tableCell of row.cells) {
        expect(tableCell.properties?.child("noFill").exists()).toBe(true);
        expect(tableCell.properties?.child("lnB").child("noFill").exists()).toBe(true);
      }
    }
  });

  it("sets which parts of the table style show", async () => {
    const pres = await openEditable([twoByTwo()]);
    const slideId = pres.slides[0].id;
    await applyEdit(pres, {
      type: "setTableOptions",
      slideId,
      nodeId: "2",
      options: { headerRow: true, bandedRows: true },
    });
    expect(tableOf(pres).properties?.attr("firstRow")).toBe("1");
    expect(tableOf(pres).properties?.attr("bandRow")).toBe("1");

    await applyEdit(pres, {
      type: "setTableOptions",
      slideId,
      nodeId: "2",
      options: { headerRow: false },
    });
    expect(tableOf(pres).properties?.attr("firstRow")).toBeUndefined();
    expect(tableOf(pres).properties?.attr("bandRow")).toBe("1");
  });
});
