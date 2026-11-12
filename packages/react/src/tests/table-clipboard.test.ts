import { describe, expect, it } from "vitest";

import { readCellGrid, writeCellGrid } from "../selection/table-clipboard";

describe("readCellGrid", () => {
  it("reads a spreadsheet's HTML table, including a single column", () => {
    const html =
      "<google-sheets-html-origin><table><tbody><tr><td>A</td></tr><tr><td>B<br>C</td></tr></tbody></table>";
    expect(readCellGrid("A\nB C", html)).toEqual([["A"], ["B\nC"]]);
  });

  it("reads tab-separated text with quoted values", () => {
    expect(readCellGrid('A\t"B\tB"\r\n"C\nC"\t"say ""hi"""\r\n', "")).toEqual([
      ["A", "B\tB"],
      ["C\nC", 'say "hi"'],
    ]);
  });

  it("leaves a single value and plain text to the cell", () => {
    expect(
      readCellGrid("one", "<table><tr><td>one</td></tr></table>"),
    ).toBeNull();
    expect(readCellGrid("line one\nline two", "")).toBeNull();
    expect(readCellGrid("Name:\tJohn", "")).toBeNull();
  });

  it("reads only a table's own rows, and spreads merged columns", () => {
    const html =
      '<table><tr><td colspan="2">A</td><td><table><tr><td>x</td></tr></table></td></tr><tr><td>B</td><td>C</td><td>D</td></tr></table>';
    expect(readCellGrid("", html)).toEqual([
      ["A", "", "x"],
      ["B", "C", "D"],
    ]);
  });
});

describe("writeCellGrid", () => {
  it("quotes values a tab-separated cell can't hold, and writes an HTML table", () => {
    const { text, html } = writeCellGrid([
      ["A", "B\nC"],
      ['say "hi"', "<b>"],
    ]);
    expect(text).toBe('A\t"B\nC"\n"say ""hi"""\t<b>');
    expect(html).toBe(
      '<table><tr><td>A</td><td>B<br>C</td></tr><tr><td>say "hi"</td><td>&lt;b&gt;</td></tr></table>',
    );
    expect(readCellGrid(text, "")).toEqual([
      ["A", "B\nC"],
      ['say "hi"', "<b>"],
    ]);
    expect(readCellGrid("", html)).toEqual([
      ["A", "B\nC"],
      ['say "hi"', "<b>"],
    ]);
  });
});
