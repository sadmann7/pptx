export function readCellGrid(text: string, html: string): string[][] | null {
  const grid = (html && readHtmlGrid(html)) || readTsvGrid(text);
  return grid && (grid.length > 1 || grid[0].length > 1) ? grid : null;
}

export function writeCellGrid(rows: string[][]): {
  text: string;
  html: string;
} {
  const text = rows
    .map((row) =>
      row
        .map((value) =>
          /[\t\n"]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value,
        )
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
  const table = new DOMParser()
    .parseFromString(html, "text/html")
    .querySelector("table");
  if (!table) return null;
  const ownRows = table.querySelectorAll<HTMLTableRowElement>(
    ":scope > tr, :scope > thead > tr, :scope > tbody > tr, :scope > tfoot > tr",
  );
  const rows = Array.from(ownRows).map((tr) =>
    Array.from(
      tr.querySelectorAll<HTMLTableCellElement>(":scope > td, :scope > th"),
    ).flatMap((cell) => {
      for (const br of Array.from(cell.querySelectorAll("br")))
        br.replaceWith("\n");
      const value = (cell.textContent ?? "").replace(/\r/g, "").trim();
      return [value, ...Array.from({ length: cell.colSpan - 1 }, () => "")];
    }),
  );
  return rows.length > 0 && rows[0].length > 0 ? rows : null;
}

function readTsvGrid(text: string): string[][] | null {
  // A single line with a tab in it is text, not a grid.
  if (!text.includes("\t") || !/\r?\n./.test(text)) return null;
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
  if (rows.length > 1 && rows[rows.length - 1].every((cell) => cell === ""))
    rows.pop();
  return rows;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
