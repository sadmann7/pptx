/**
 * @license
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * @remarks
 * Derived from mtx-decompressor v1.4.2
 * (https://github.com/ChristopherVR/mtx-decompressor, © ChristopherVR, MPL-2.0).
 * Optimized for `@diceui/pptx-core`:
 * - Triplet encodings as flat typed arrays (no object field loads)
 * - Bulk table copies instead of per-byte reads
 * - Reusable per-glyph scratch buffers instead of per-glyph allocations
 *
 * Output is byte-identical to the original.
 */

import { MtxStream } from "./stream";

// ── TrueType constants ──────────────────────────────────────────────

const FLG_ON_CURVE = 1;
const FLG_X_SHORT = 2;
const FLG_Y_SHORT = 4;
const FLG_X_SAME = 16;
const FLG_Y_SAME = 32;

const NPUSHB = 64;
const NPUSHW = 65;
const PUSHB = 176;
const PUSHW = 184;

const ARG_1_AND_2_ARE_WORDS = 1;
const HAVE_SCALE = 8;
const MORE_COMPONENTS = 32;
const HAVE_XY_SCALE = 64;
const HAVE_2_BY_2 = 128;
const HAVE_INSTRUCTIONS = 256;

const INT16_MIN = -32768;
const INT16_MAX = 32767;

// ── Triplet encodings (flat structure-of-arrays) ────────────────────
// Index = flag & 127. Generated to match the reference table exactly.

const TRIPLET_COUNT = 128;
const T_BYTE_COUNT = new Uint8Array(TRIPLET_COUNT);
const T_X_BITS = new Uint8Array(TRIPLET_COUNT);
const T_Y_BITS = new Uint8Array(TRIPLET_COUNT);
const T_DELTA_X = new Int32Array(TRIPLET_COUNT);
const T_DELTA_Y = new Int32Array(TRIPLET_COUNT);
const T_X_SIGN = new Int8Array(TRIPLET_COUNT);
const T_Y_SIGN = new Int8Array(TRIPLET_COUNT);

{
  let i = 0;
  const put = (
    byteCount: number,
    xBits: number,
    yBits: number,
    deltaX: number,
    deltaY: number,
    xSign: number,
    ySign: number,
  ): void => {
    T_BYTE_COUNT[i] = byteCount;
    T_X_BITS[i] = xBits;
    T_Y_BITS[i] = yBits;
    T_DELTA_X[i] = deltaX;
    T_DELTA_Y[i] = deltaY;
    T_X_SIGN[i] = xSign;
    T_Y_SIGN[i] = ySign;
    i++;
  };

  // 0-9: Y axis only, deltas 0..1024 step 256, alternating sign
  for (const delta of [0, 256, 512, 768, 1024]) {
    put(2, 0, 8, 0, delta, 0, -1);
    put(2, 0, 8, 0, delta, 0, 1);
  }
  // 10-19: X axis only
  for (const delta of [0, 256, 512, 768, 1024]) {
    put(2, 8, 0, delta, 0, -1, 0);
    put(2, 8, 0, delta, 0, 1, 0);
  }
  // 20-83: 4-bit X + 4-bit Y, deltas {1,17,33,49} x {1,17,33,49}, 4 sign combos
  for (const dx of [1, 17, 33, 49]) {
    for (const dy of [1, 17, 33, 49]) {
      put(2, 4, 4, dx, dy, -1, -1);
      put(2, 4, 4, dx, dy, 1, -1);
      put(2, 4, 4, dx, dy, -1, 1);
      put(2, 4, 4, dx, dy, 1, 1);
    }
  }
  // 84-119: 8-bit X + 8-bit Y, deltas {1,257,513} x {1,257,513}
  for (const dx of [1, 257, 513]) {
    for (const dy of [1, 257, 513]) {
      put(3, 8, 8, dx, dy, -1, -1);
      put(3, 8, 8, dx, dy, 1, -1);
      put(3, 8, 8, dx, dy, -1, 1);
      put(3, 8, 8, dx, dy, 1, 1);
    }
  }
  // 120-123: 12-bit X + 12-bit Y
  put(4, 12, 12, 0, 0, -1, -1);
  put(4, 12, 12, 0, 0, 1, -1);
  put(4, 12, 12, 0, 0, -1, 1);
  put(4, 12, 12, 0, 0, 1, 1);
  // 124-127: 16-bit X + 16-bit Y
  put(5, 16, 16, 0, 0, -1, -1);
  put(5, 16, 16, 0, 0, 1, -1);
  put(5, 16, 16, 0, 0, -1, 1);
  put(5, 16, 16, 0, 0, 1, 1);
}

// ── Table container ─────────────────────────────────────────────────

export interface CtfTable {
  tag: string;
  offset: number;
  bufSize: number;
  buf: Uint8Array;
  checksum: number;
}

export interface CtfContainer {
  tables: CtfTable[];
}

// ── Variable-length integer readers ─────────────────────────────────

function toInt16(v: number): number {
  v &= 65535;
  return v >= 32768 ? v - 65536 : v;
}

function read255UShort(s: MtxStream): number {
  const code = s.readU8();
  if (code === 253) return s.readU16();
  if (code === 255) return 253 + s.readU8();
  if (code === 254) return 506 + s.readU8();
  return code;
}

function read255Short(s: MtxStream): number {
  let sign = 1;
  let code = s.readU8();
  if (code === 253) return s.readS16();
  if (code === 250) {
    sign = -1;
    code = s.readU8();
  }
  let value: number;
  if (code === 255) {
    value = 250 + s.readU8();
  } else if (code === 254) {
    value = 500 + s.readU8();
  } else {
    value = code;
  }
  return value * sign;
}

// ── CVT ─────────────────────────────────────────────────────────────

function unpackCVT(table: CtfTable, sIn: MtxStream): void {
  sIn.seekAbsolute(table.offset);
  const tableLength = sIn.readU16();
  const numEntries = tableLength >>> 1;
  const out = new MtxStream(null, 0);
  out.reserve(tableLength);
  let lastValue = 0;
  for (let i = 0; i < numEntries; i++) {
    const code = sIn.readU8();
    let val: number;
    if (code >= 248) {
      val = 238 * (code - 247) + sIn.readU8();
    } else if (code >= 239) {
      val = -(238 * (code - 239) + sIn.readU8());
    } else if (code === 238) {
      val = sIn.readS16();
    } else {
      val = code;
    }
    lastValue = toInt16(lastValue + val);
    out.writeS16(lastValue);
  }
  table.buf = out.toUint8Array();
  table.bufSize = table.buf.length;
}

// ── Instruction push decoding ───────────────────────────────────────

function decodePushInstructions(sIn: MtxStream, sOut: MtxStream, pushCount: number): void {
  if (pushCount === 0) return;

  const data: number[] = [];
  let remaining = pushCount;
  let isShort = false;
  const runValues: number[] = [];

  const flush = (): void => {
    const count = runValues.length;
    if (count === 0) return;
    if (isShort) {
      if (count < 8) {
        sOut.writeU8(PUSHW + (count - 1));
      } else {
        sOut.writeU8(NPUSHW);
        sOut.writeU8(count);
      }
      for (let i = 0; i < count; i++) sOut.writeS16(runValues[i]);
    } else {
      if (count < 8) {
        sOut.writeU8(PUSHB + (count - 1));
      } else {
        sOut.writeU8(NPUSHB);
        sOut.writeU8(count);
      }
      for (let i = 0; i < count; i++) sOut.writeU8(runValues[i] & 255);
    }
    runValues.length = 0;
  };

  const put = (v: number): void => {
    data.push(v);
    const needsShort = v < 0 || v > 255;
    if (runValues.length > 0 && needsShort !== isShort) {
      flush();
    }
    if (runValues.length === 0) {
      isShort = needsShort;
    }
    runValues.push(v);
    if (runValues.length >= 255) {
      flush();
    }
  };

  while (remaining > 0) {
    const code = sIn.peekU8();
    if (code === 251 && remaining >= 3 && data.length >= 2) {
      sIn.readU8();
      const prev = data[data.length - 2];
      put(prev);
      put(read255Short(sIn));
      put(prev);
      remaining -= 3;
    } else if (code === 252 && remaining >= 5 && data.length >= 2) {
      sIn.readU8();
      const prev = data[data.length - 2];
      put(prev);
      put(read255Short(sIn));
      put(prev);
      put(read255Short(sIn));
      put(prev);
      remaining -= 5;
    } else {
      put(read255Short(sIn));
      remaining -= 1;
    }
  }
  flush();
}

// ── Glyph decoding ──────────────────────────────────────────────────

function makeGlyphFlags(x: number, y: number, onCurve: boolean, firstTime: boolean): number {
  let flags = 0;
  if (onCurve) flags |= FLG_ON_CURVE;
  if (!firstTime && x === 0) {
    flags |= FLG_X_SAME;
  } else if (x > -256 && x < 0) {
    flags |= FLG_X_SHORT;
  } else if (x >= 0 && x < 256) {
    flags |= FLG_X_SHORT | FLG_X_SAME;
  }
  if (!firstTime && y === 0) {
    flags |= FLG_Y_SAME;
  } else if (y > -256 && y < 0) {
    flags |= FLG_Y_SHORT;
  } else if (y >= 0 && y < 256) {
    flags |= FLG_Y_SHORT | FLG_Y_SAME;
  }
  return flags;
}

/** Reusable per-decode scratch to avoid per-glyph typed-array churn. */
interface GlyphScratch {
  flagBytes: Uint8Array;
  xDeltas: Int16Array;
  yDeltas: Int16Array;
  onCurve: Uint8Array;
}

function ensureScratch(scratch: GlyphScratch, totalPoints: number): void {
  if (scratch.flagBytes.length >= totalPoints) return;
  const cap = Math.max(totalPoints, scratch.flagBytes.length * 2 || 256);
  scratch.flagBytes = new Uint8Array(cap);
  scratch.xDeltas = new Int16Array(cap);
  scratch.yDeltas = new Int16Array(cap);
  scratch.onCurve = new Uint8Array(cap);
}

function decodeSimpleGlyph(
  numContours: number,
  streams: MtxStream[],
  out: MtxStream,
  calcBBox: boolean,
  minX: number,
  minY: number,
  maxX: number,
  maxY: number,
  scratch: GlyphScratch,
): void {
  if (numContours === 0) return;

  const sGlyph = streams[0];
  out.writeS16(numContours);
  const bboxPos = out.pos;
  if (calcBBox) {
    minX = INT16_MAX;
    minY = INT16_MAX;
    maxX = INT16_MIN;
    maxY = INT16_MIN;
    out.writeS16(0);
    out.writeS16(0);
    out.writeS16(0);
    out.writeS16(0);
  } else {
    out.writeS16(minX);
    out.writeS16(minY);
    out.writeS16(maxX);
    out.writeS16(maxY);
  }

  let totalPoints = 0;
  for (let c = 0; c < numContours; c++) {
    if (c === 0) totalPoints = 1;
    totalPoints += read255UShort(sGlyph);
    out.writeU16(totalPoints - 1);
  }

  ensureScratch(scratch, totalPoints);
  const { flagBytes, xDeltas, yDeltas, onCurve } = scratch;

  for (let i = 0; i < totalPoints; i++) {
    flagBytes[i] = sGlyph.readU8();
  }

  let cumulativeX = 0;
  let cumulativeY = 0;
  for (let i = 0; i < totalPoints; i++) {
    const flag = flagBytes[i];
    onCurve[i] = flag & 128 ? 0 : 1;
    const enc = flag & 127;
    let dx = sGlyph.readNBits(T_X_BITS[enc]) + T_DELTA_X[enc];
    let dy = sGlyph.readNBits(T_Y_BITS[enc]) + T_DELTA_Y[enc];
    const xs = T_X_SIGN[enc];
    const ys = T_Y_SIGN[enc];
    if (xs !== 0) dx *= xs;
    if (ys !== 0) dy *= ys;
    xDeltas[i] = dx;
    yDeltas[i] = dy;
    cumulativeX += dx;
    cumulativeY += dy;
    if (calcBBox) {
      if (cumulativeX < minX) minX = cumulativeX;
      if (cumulativeX > maxX) maxX = cumulativeX;
      if (cumulativeY < minY) minY = cumulativeY;
      if (cumulativeY > maxY) maxY = cumulativeY;
    }
  }

  const codeSizeLocation = out.pos;
  out.writeU16(0);
  const pushCount = read255UShort(sGlyph);
  decodePushInstructions(streams[1], out, pushCount);
  const codeSize = read255UShort(sGlyph);
  if (codeSize > 0) {
    streams[2].copyTo(out, codeSize);
  }
  const unpackedCodeSize = out.pos - (codeSizeLocation + 2);
  const savedPos = out.pos;
  out.seekAbsolute(codeSizeLocation);
  out.writeU16(unpackedCodeSize);
  out.seekAbsolute(savedPos);

  for (let i = 0; i < totalPoints; i++) {
    out.writeU8(makeGlyphFlags(xDeltas[i], yDeltas[i], onCurve[i] !== 0, i === 0));
  }
  for (let i = 0; i < totalPoints; i++) {
    const x = xDeltas[i];
    if (i === 0 || x !== 0) {
      const absX = x < 0 ? -x : x;
      if (absX < 256) {
        out.writeU8(absX);
      } else {
        out.writeS16(x);
      }
    }
  }
  for (let i = 0; i < totalPoints; i++) {
    const y = yDeltas[i];
    if (i === 0 || y !== 0) {
      const absY = y < 0 ? -y : y;
      if (absY < 256) {
        out.writeU8(absY);
      } else {
        out.writeS16(y);
      }
    }
  }

  if (calcBBox) {
    const endPos = out.pos;
    out.seekAbsolute(bboxPos);
    out.writeS16(minX);
    out.writeS16(minY);
    out.writeS16(maxX);
    out.writeS16(maxY);
    out.seekAbsolute(endPos);
  }
}

function decodeCompositeGlyph(streams: MtxStream[], out: MtxStream): void {
  const sGlyph = streams[0];
  out.writeS16(-1);
  out.writeS16(sGlyph.readS16());
  out.writeS16(sGlyph.readS16());
  out.writeS16(sGlyph.readS16());
  out.writeS16(sGlyph.readS16());
  let flags = 0;
  do {
    flags = sGlyph.readU16();
    const glyphIndex = sGlyph.readU16();
    out.writeU16(flags);
    out.writeU16(glyphIndex);
    sGlyph.copyTo(out, flags & ARG_1_AND_2_ARE_WORDS ? 4 : 2);
    let transformBytes = 0;
    if (flags & HAVE_2_BY_2) {
      transformBytes = 8;
    } else if (flags & HAVE_XY_SCALE) {
      transformBytes = 4;
    } else if (flags & HAVE_SCALE) {
      transformBytes = 2;
    }
    if (transformBytes > 0) {
      sGlyph.copyTo(out, transformBytes);
    }
  } while (flags & MORE_COMPONENTS);

  if (flags & HAVE_INSTRUCTIONS) {
    const numInstrPos = out.pos;
    out.writeU16(0);
    const pushCount = read255UShort(sGlyph);
    decodePushInstructions(streams[1], out, pushCount);
    const codeSize = read255UShort(sGlyph);
    if (codeSize > 0) {
      streams[2].copyTo(out, codeSize);
    }
    const numInstr = out.pos - (numInstrPos + 2);
    const savedPos = out.pos;
    out.seekAbsolute(numInstrPos);
    out.writeU16(numInstr);
    out.seekAbsolute(savedPos);
  }
}

function decodeGlyph(streams: MtxStream[], out: MtxStream, scratch: GlyphScratch): void {
  const numContours = streams[0].readS16();
  if (numContours < 0) {
    decodeCompositeGlyph(streams, out);
  } else if (numContours === 32767) {
    const actualContours = streams[0].readS16();
    const xMin = streams[0].readS16();
    const yMin = streams[0].readS16();
    const xMax = streams[0].readS16();
    const yMax = streams[0].readS16();
    decodeSimpleGlyph(actualContours, streams, out, false, xMin, yMin, xMax, yMax, scratch);
  } else {
    decodeSimpleGlyph(numContours, streams, out, true, 0, 0, 0, 0, scratch);
  }
}

// ── glyf/loca reconstruction ────────────────────────────────────────

interface HeadData {
  indexToLocFormat: number;
}

interface MaxpData {
  numGlyphs: number;
  maxPoints: number;
  maxContours: number;
  maxSizeOfInstructions: number;
  maxComponentElements: number;
}

function populateGlyfAndLoca(
  glyf: CtfTable,
  loca: CtfTable,
  headData: HeadData,
  maxpData: MaxpData,
  streams: MtxStream[],
): void {
  const numGlyphs = maxpData.numGlyphs;
  streams[0].seekAbsolute(glyf.offset);
  streams[1].seekAbsolute(0);
  streams[2].seekAbsolute(0);

  const maxGlyphSize =
    5 * 2 +
    2 * maxpData.maxContours +
    2 +
    maxpData.maxSizeOfInstructions +
    256 +
    5 * maxpData.maxPoints +
    4 * maxpData.maxComponentElements * 6 +
    256;

  const outStream = new MtxStream(null, 0);
  outStream.reserve(numGlyphs * 256);
  const isShortLoca = headData.indexToLocFormat === 0;
  const locaStream = new MtxStream(null, 0);
  locaStream.reserve((numGlyphs + 1) * (isShortLoca ? 2 : 4));
  if (isShortLoca) {
    locaStream.writeU16(0);
  } else {
    locaStream.writeU32(0);
  }

  const scratch: GlyphScratch = {
    flagBytes: new Uint8Array(0),
    xDeltas: new Int16Array(0),
    yDeltas: new Int16Array(0),
    onCurve: new Uint8Array(0),
  };

  for (let i = 0; i < numGlyphs; i++) {
    // Geometric growth: the original re-reserved exactly pos+maxGlyphSize per
    // glyph, forcing O(n²) copies on fonts whose glyf outgrows the estimate.
    outStream.reserveGrow(outStream.pos + maxGlyphSize);
    decodeGlyph(streams, outStream, scratch);
    if (outStream.pos & 1) {
      outStream.writeU8(0);
    }
    if (isShortLoca) {
      locaStream.writeU16(outStream.pos >>> 1);
    } else {
      locaStream.writeU32(outStream.pos);
    }
  }

  glyf.buf = outStream.toUint8Array();
  glyf.bufSize = glyf.buf.length;
  loca.buf = locaStream.toUint8Array();
  loca.bufSize = loca.buf.length;
}

function parseHead(table: CtfTable): HeadData {
  const s = new MtxStream(table.buf, table.bufSize);
  s.seekAbsolute(50);
  return { indexToLocFormat: s.readS16() };
}

function parseMaxp(table: CtfTable): MaxpData {
  const s = new MtxStream(table.buf, table.bufSize);
  const version = s.readU32();
  const numGlyphs = s.readU16();
  let maxPoints = 0;
  let maxContours = 0;
  let maxSizeOfInstructions = 0;
  let maxComponentElements = 0;
  if (version === 65536) {
    maxPoints = s.readU16();
    maxContours = s.readU16();
    s.seekRelative(16);
    maxSizeOfInstructions = s.readU16();
    maxComponentElements = s.readU16();
  }
  return { numGlyphs, maxPoints, maxContours, maxSizeOfInstructions, maxComponentElements };
}

// ── CTF container parsing ───────────────────────────────────────────

export function parseCTF(streams: MtxStream[]): CtfContainer {
  const s0 = streams[0];
  s0.readU32();
  const numTables = s0.readU16();
  s0.seekRelative(6);

  const tables: CtfTable[] = [];
  let glyfIdx = -1;
  let locaIdx = -1;
  let maxpIdx = -1;
  let headIdx = -1;

  for (let i = 0; i < numTables; i++) {
    const t0 = s0.readU8();
    const t1 = s0.readU8();
    const t2 = s0.readU8();
    const t3 = s0.readU8();
    const tag = String.fromCharCode(t0, t1, t2, t3);
    if (tag === "hdmx" || tag === "VDMX") {
      s0.seekRelative(12);
      continue;
    }
    s0.seekRelative(4);
    const offset = s0.readU32();
    const size = s0.readU32();
    const table: CtfTable = { tag, offset, bufSize: size, buf: new Uint8Array(0), checksum: 0 };
    const idx = tables.length;
    tables.push(table);
    if (tag === "glyf") glyfIdx = idx;
    else if (tag === "loca") locaIdx = idx;
    else if (tag === "maxp") maxpIdx = idx;
    else if (tag === "head") headIdx = idx;
  }

  for (const table of tables) {
    if (table.tag === "glyf" || table.tag === "loca") continue;
    if (table.tag === "cvt ") {
      unpackCVT(table, s0);
      continue;
    }
    s0.seekAbsolute(table.offset);
    s0.ensureRead(table.bufSize);
    // Bulk copy (the original read byte-by-byte through readU8).
    table.buf = s0.buf.slice(s0.pos, s0.pos + table.bufSize);
    s0.pos += table.bufSize;
    if (table.tag === "head") {
      table.buf[8] = 0;
      table.buf[9] = 0;
      table.buf[10] = 0;
      table.buf[11] = 0;
    }
  }

  let headData: HeadData = { indexToLocFormat: 0 };
  if (headIdx >= 0) headData = parseHead(tables[headIdx]);

  let maxpData: MaxpData = {
    numGlyphs: 0,
    maxPoints: 0,
    maxContours: 0,
    maxSizeOfInstructions: 0,
    maxComponentElements: 0,
  };
  if (maxpIdx >= 0) maxpData = parseMaxp(tables[maxpIdx]);

  if (glyfIdx >= 0) {
    if (locaIdx < 0) {
      locaIdx = tables.length;
      tables.push({ tag: "loca", offset: 0, bufSize: 0, buf: new Uint8Array(0), checksum: 0 });
    }
    populateGlyfAndLoca(tables[glyfIdx], tables[locaIdx], headData, maxpData, streams);
  }

  return { tables };
}
