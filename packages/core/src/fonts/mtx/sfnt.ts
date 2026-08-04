import { Writer } from "./binary";
import { fail } from "./error";

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

export function tagAt(data: Uint8Array, offset: number): string {
  if (offset < 0 || offset + 4 > data.length) fail("BOUNDS", "Tag outside input", offset);
  return String.fromCharCode(
    data[offset]!,
    data[offset + 1]!,
    data[offset + 2]!,
    data[offset + 3]!,
  );
}

export function tagBytes(tag: string): Uint8Array {
  if (tag.length !== 4) throw new TypeError("SFNT tags must contain four characters");
  return Uint8Array.from([
    tag.charCodeAt(0),
    tag.charCodeAt(1),
    tag.charCodeAt(2),
    tag.charCodeAt(3),
  ]);
}

/** Sum of the data as big-endian uint32 words, zero-padded to a word boundary. */
export function checksum(data: Uint8Array): number {
  let sum = 0;
  for (let i = 0; i < data.length; i += 4) {
    const word =
      (data[i] ?? 0) * 0x1000000 +
      ((data[i + 1] ?? 0) << 16) +
      ((data[i + 2] ?? 0) << 8) +
      (data[i + 3] ?? 0);
    sum = (sum + word) >>> 0;
  }
  return sum;
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
