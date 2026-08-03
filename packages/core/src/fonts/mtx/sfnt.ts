import { checksum, tagBytes, Writer } from "./binary";
import { fail } from "./errors";

export interface SfntTable {
  tag: string;
  data: Uint8Array;
  checksum?: number;
  offset?: number;
}

export interface SfntContainer {
  sfntVersion: number;
  tables: SfntTable[];
  droppedTables?: string[];
}

function highestPowerOfTwo(value: number): number {
  let result = 1;
  while (result * 2 <= value) result *= 2;
  return result;
}

export function buildSfnt(container: SfntContainer, maxBytes: number): Uint8Array {
  // OpenType requires directory entries in ascending order by raw tag bytes.
  // Locale-aware collation groups by letter regardless of case and would place
  // "cvt " before "OS/2".
  const tables = [...container.tables].sort((a, b) => (a.tag < b.tag ? -1 : a.tag > b.tag ? 1 : 0));
  const count = tables.length;
  if (count === 0) fail("INVALID_SFNT", "Cannot build an SFNT without tables");
  const directorySize = 12 + count * 16;
  let totalSize = directorySize;
  for (const table of tables) totalSize += (table.data.length + 3) & ~3;
  if (totalSize > maxBytes) fail("LIMIT_EXCEEDED", `SFNT output requires ${totalSize} bytes`);

  const writer = new Writer(totalSize);
  writer.u32be(container.sfntVersion);
  writer.u16be(count);
  const power = highestPowerOfTwo(count);
  writer.u16be(power * 16);
  writer.u16be(Math.log2(power));
  writer.u16be(count * 16 - power * 16);

  let offset = directorySize;
  for (const table of tables) {
    table.offset = offset;
    table.checksum = checksum(table.data);
    writer.bytes(tagBytes(table.tag));
    writer.u32be(table.checksum);
    writer.u32be(offset);
    writer.u32be(table.data.length);
    offset += (table.data.length + 3) & ~3;
  }
  for (const table of tables) {
    writer.bytes(table.data);
    writer.align(4);
  }

  const font = writer.finish();
  const head = tables.find((table) => table.tag === "head");
  if (!head || head.offset === undefined || head.data.length < 12)
    fail("INVALID_SFNT", "Missing or malformed head table");
  const adjustment = (0xb1b0afba - checksum(font)) >>> 0;
  new DataView(font.buffer, font.byteOffset, font.byteLength).setUint32(
    head.offset + 8,
    adjustment,
    false,
  );
  return font;
}

export function validateSfnt(font: Uint8Array): void {
  if (font.length < 12) fail("INVALID_SFNT", "SFNT is truncated");
  const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
  const count = view.getUint16(4, false);
  if (12 + count * 16 > font.length) fail("INVALID_SFNT", "SFNT directory is truncated");
  let hasHead = false;
  for (let i = 0; i < count; i++) {
    const record = 12 + i * 16;
    const tag = String.fromCharCode(
      font[record]!,
      font[record + 1]!,
      font[record + 2]!,
      font[record + 3]!,
    );
    const offset = view.getUint32(record + 8, false);
    const length = view.getUint32(record + 12, false);
    if (offset + length > font.length) fail("INVALID_SFNT", `Table ${tag} lies outside the font`);
    if (tag === "head") hasHead = true;
  }
  if (!hasHead) fail("INVALID_SFNT", "SFNT has no head table");
  if (checksum(font) !== 0xb1b0afba) fail("INVALID_SFNT", "Invalid whole-font checksum");
}
