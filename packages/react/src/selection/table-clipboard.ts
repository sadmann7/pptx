/**
 * Reads the cells a spreadsheet put on the clipboard: its HTML table when there
 * is one, else tab-separated text. `null` when the clipboard holds no grid of
 * more than one cell, so the text pastes into the cell as usual.
 */
export function readCellGrid(text: string, html: string): string[][] | null {
  const grid = (html && readHtmlGrid(html)) || readTsvGrid(text);
  return grid && (grid.length > 1 || grid[0].length > 1) ? grid : null;
}

/** Cells as spreadsheets copy them: tab-separated text and an HTML table. */
export function writeCellGrid(rows: string[][]): { text: string; html: string } {
  const text = rows
    .map((row) =>
      row
        .map((value) => (/[\t\n"]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value))
        .join("\t"),
    )
    .join("\n");
  const html = `<table>${rows
    .map(
      (row) =>
        `<tr>${row.map((value) => `<td>${escapeHtml(value).replace(/\n/g, "<br>")}</td>`).join("")}</tr>`,
    )
    .join("")}</table>`;
  return { text, html };
}

function readHtmlGrid(html: string): string[][] | null {
  const table = new DOMParser().parseFromString(html, "text/html").querySelector("table");
  if (!table) return null;
  const rows = Array.from(table.querySelectorAll("tr")).map((tr) =>
    Array.from(tr.querySelectorAll("td, th")).map((cell) => {
      for (const br of Array.from(cell.querySelectorAll("br"))) br.replaceWith("\n");
      return (cell.textContent ?? "").replace(/\r/g, "").trim();
    }),
  );
  return rows.length > 0 && rows[0].length > 0 ? rows : null;
}

/** Tab-separated text, with quoted values that can hold tabs, line breaks, and quotes. */
function readTsvGrid(text: string): string[][] | null {
  if (!text.includes("\t")) return null;
  const rows: string[][] = [[]];
  let value = "";
  let quoted = false;
  for (let index = 0; index < text.length; index++) {
    const char = text[index];
    if (quoted) {
      if (char === '"' && text[index + 1] === '"') {
        value += '"';
        index++;
      } else if (char === '"') {
        quoted = false;
      } else {
        value += char;
      }
    } else if (char === '"' && value === "") {
      quoted = true;
    } else if (char === "\t") {
      rows[rows.length - 1].push(value);
      value = "";
    } else if (char === "\n" || char === "\r") {
      if (char === "\r" && text[index + 1] === "\n") index++;
      rows[rows.length - 1].push(value);
      value = "";
      rows.push([]);
    } else {
      value += char;
    }
  }
  rows[rows.length - 1].push(value);
  // A trailing line break leaves an empty last row.
  if (rows.length > 1 && rows[rows.length - 1].every((cell) => cell === "")) rows.pop();
  return rows;
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
