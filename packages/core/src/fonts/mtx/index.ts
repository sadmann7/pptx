/**
 * @license
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * Derived from mtx-decompressor v1.4.2
 * (https://github.com/ChristopherVR/mtx-decompressor, © ChristopherVR,
 * MPL-2.0).
 *
 * @remarks
 * Optimized for @diceui/pptx-core; output is byte-identical
 * to the original.
 */

import type { CtfContainer, CtfTable } from "./ctf";
import { parseCTF } from "./ctf";
import { lzcompDecompress } from "./lzcomp";
import { MtxStream } from "./stream";

const ENCRYPTION_KEY = 80;

export interface DecompressMtxOptions {
  /** XOR-decrypt the container before unpacking. */
  encrypted?: boolean;
  /** Whether the payload is MTX-compressed (as opposed to a raw font). */
  compressed?: boolean;
}

// ── SFNT container serialization ────────────────────────────────────

function lgFloor(n: number): number {
  let ret = 0;
  while (n > 1) {
    n >>= 1;
    ret++;
  }
  return ret;
}

function writeTableCheckingSum(table: CtfTable, out: MtxStream): number {
  table.offset = out.pos;
  const data = table.buf;
  const len = table.bufSize;
  const fullWords = len >>> 2;
  const remainder = len & 3;

  out.ensureWrite(fullWords * 4 + (remainder > 0 ? 4 : 0));
  const buf = out.buf;
  let pos = out.pos;
  let checksum = 0;
  for (let i = 0; i < fullWords; i++) {
    const off = i * 4;
    const word =
      ((data[off] << 24) | (data[off + 1] << 16) | (data[off + 2] << 8) | data[off + 3]) >>> 0;
    checksum = (checksum + word) >>> 0;
    buf[pos++] = (word >>> 24) & 255;
    buf[pos++] = (word >>> 16) & 255;
    buf[pos++] = (word >>> 8) & 255;
    buf[pos++] = word & 255;
  }
  if (remainder > 0) {
    let word = 0;
    for (let j = 0; j < remainder; j++) {
      word |= data[fullWords * 4 + j] << (24 - j * 8);
    }
    word >>>= 0;
    checksum = (checksum + word) >>> 0;
    buf[pos++] = (word >>> 24) & 255;
    buf[pos++] = (word >>> 16) & 255;
    buf[pos++] = (word >>> 8) & 255;
    buf[pos++] = word & 255;
  }
  out.pos = pos;
  table.checksum = checksum;
  return checksum;
}

function dumpContainer(ctr: CtfContainer): Uint8Array {
  const numTables = ctr.tables.length;
  const dirSize = 16 * numTables;
  let tableDataSize = 0;
  for (const table of ctr.tables) {
    tableDataSize += (table.bufSize + 3) & ~3;
  }
  const out = new MtxStream(new Uint8Array(12 + dirSize + tableDataSize), 0);

  // Offset table
  const searchRange = (1 << lgFloor(numTables)) * 16;
  out.writeU32(65536);
  out.writeU16(numTables);
  out.writeU16(searchRange);
  out.writeU16(lgFloor(numTables));
  out.writeU16(numTables * 16 - searchRange);

  const dirOffset = out.pos;
  out.ensureWrite(dirSize);
  out.pos += dirSize;

  let totalChecksum = 0;
  let headTable: CtfTable | undefined;
  for (const table of ctr.tables) {
    totalChecksum = (totalChecksum + writeTableCheckingSum(table, out)) >>> 0;
    if (table.tag === "head") headTable = table;
  }

  const afterTables = out.pos;
  out.pos = dirOffset;
  for (const table of ctr.tables) {
    out.writeU8(table.tag.charCodeAt(0));
    out.writeU8(table.tag.charCodeAt(1));
    out.writeU8(table.tag.charCodeAt(2));
    out.writeU8(table.tag.charCodeAt(3));
    out.writeU32(table.checksum);
    out.writeU32(table.offset);
    out.writeU32(table.bufSize);
  }

  const beginningLen = 12 + dirSize;
  const buf = out.buf;
  let beginningChecksum = 0;
  for (let off = 0; off + 4 <= beginningLen; off += 4) {
    const word =
      ((buf[off] << 24) | (buf[off + 1] << 16) | (buf[off + 2] << 8) | buf[off + 3]) >>> 0;
    beginningChecksum = (beginningChecksum + word) >>> 0;
  }
  totalChecksum = (totalChecksum + beginningChecksum) >>> 0;
  const finalChecksum = (2981146554 - totalChecksum) >>> 0;
  if (headTable) {
    const adjOffset = headTable.offset + 8;
    buf[adjOffset] = (finalChecksum >>> 24) & 255;
    buf[adjOffset + 1] = (finalChecksum >>> 16) & 255;
    buf[adjOffset + 2] = (finalChecksum >>> 8) & 255;
    buf[adjOffset + 3] = finalChecksum & 255;
  }

  return buf.subarray(0, afterTables);
}

// ── MTX unpacking ───────────────────────────────────────────────────

function unpackMtx(data: Uint8Array, size: number): Uint8Array[] {
  if (size < 10 || data.length < 10) {
    throw new Error("MTX data too small: header requires at least 10 bytes");
  }
  const versionMagic = data[0];
  const offset2 = (data[4] << 16) | (data[5] << 8) | data[6];
  const offset3 = (data[7] << 16) | (data[8] << 8) | data[9];
  if (offset2 < 10 || offset3 < offset2 || offset3 > size) {
    throw new Error(
      `MTX header offsets out of bounds: offset2=${offset2}, offset3=${offset3}, size=${size}`,
    );
  }
  const offsets = [10, offset2, offset3];
  const blockSizes = [
    Math.max(0, offset2 - 10),
    Math.max(0, offset3 - offset2),
    Math.max(0, size - offset3),
  ];
  const streams: Uint8Array[] = [];
  for (let i = 0; i < 3; i++) {
    streams.push(lzcompDecompress(data.subarray(offsets[i]), blockSizes[i], versionMagic));
  }
  return streams;
}

/**
 * Decompress a MicroType Express (MTX) payload into a raw TrueType binary.
 * Drop-in replacement for the `mtx-decompressor` package's `decompressMtx`
 * with byte-identical output.
 */
export function decompressMtx(fontData: Uint8Array, options?: DecompressMtxOptions): Uint8Array {
  const encrypted = options?.encrypted ?? false;
  const compressed = options?.compressed ?? true;

  let data: Uint8Array;
  if (encrypted) {
    data = new Uint8Array(fontData.length);
    for (let i = 0; i < fontData.length; i++) {
      data[i] = fontData[i] ^ ENCRYPTION_KEY;
    }
  } else {
    data = fontData;
  }
  if (!compressed) {
    return data;
  }

  const streams = unpackMtx(data, data.length);
  const streamObjects = streams.map((buf) => new MtxStream(buf, buf.length));
  return dumpContainer(parseCTF(streamObjects));
}
