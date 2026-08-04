import { Reader, Writer } from "./binary";
import { decodeCvt } from "./cvt";
import { fail } from "./error";
import type { DecoderLimits } from "./limits";
import { readPushValuesInto, writePushInstructions } from "./push";
import { tagAt, type SfntContainer, type SfntTable } from "./sfnt";
import { decodeTripletArrays, type DecodedTriplets, type TripletScratch } from "./triplet";
import { read255UShort } from "./varint";

interface DirectoryEntry {
  tag: string;
  offset: number;
  length: number;
}

const ARG_1_AND_2_ARE_WORDS = 0x0001;
const WE_HAVE_A_SCALE = 0x0008;
const MORE_COMPONENTS = 0x0020;
const WE_HAVE_AN_X_AND_Y_SCALE = 0x0040;
const WE_HAVE_A_TWO_BY_TWO = 0x0080;
const WE_HAVE_INSTRUCTIONS = 0x0100;

function tableView(data: Uint8Array, entry: DirectoryEntry): Uint8Array {
  if (entry.offset < 0 || entry.length < 0 || entry.offset + entry.length > data.length) {
    fail("INVALID_CTF", `Table ${entry.tag} lies outside CTF stream`);
  }
  return data.subarray(entry.offset, entry.offset + entry.length);
}

/**
 * Append a glyph's instruction block to the glyf writer.
 *
 * The length prefix is reserved and patched once the PUSH burst has been
 * emitted, so the burst is measured by writing it rather than by scanning the
 * push values twice.
 */
function writeInstructions(
  rest: Reader,
  push: Reader,
  code: Reader,
  limits: DecoderLimits,
  glyphId: number,
  writer: Writer,
  scratch: GlyphScratch,
): void {
  const pushCount = read255UShort(rest);
  const codeSize = read255UShort(rest);
  if (codeSize > limits.maxInstructionsPerGlyph)
    fail("LIMIT_EXCEEDED", "Glyph instruction stream exceeds configured limit");
  if (codeSize > code.remaining) {
    fail(
      "INVALID_CTF",
      `Glyph ${glyphId} needs ${codeSize} code byte(s), only ${code.remaining} remain`,
    );
  }
  if (scratch.pushValues.length < pushCount) {
    scratch.pushValues = new Int16Array(Math.max(pushCount, scratch.pushValues.length * 2, 16));
  }
  readPushValuesInto(push, pushCount, scratch.pushValues);

  const lengthOffset = writer.length;
  writer.u16be(0);
  writePushInstructions(writer, scratch.pushValues, pushCount);
  const instructionLength = writer.length - lengthOffset - 2 + codeSize;
  if (instructionLength > 0xffff || instructionLength > limits.maxInstructionsPerGlyph) {
    fail("LIMIT_EXCEEDED", "Reconstructed glyph instructions exceed TrueType limits");
  }
  writer.patchU16be(lengthOffset, instructionLength);
  writer.bytes(code.bytes(codeSize));
}

interface GlyphScratch extends TripletScratch {
  rawFlags: Uint8Array;
  encodedFlags: Uint8Array;
  xData: Uint8Array;
  yData: Uint8Array;
  endPoints: Uint16Array;
  pushValues: Int16Array;
}

function ensureByteScratch(scratch: GlyphScratch, points: number): void {
  if (scratch.rawFlags.length >= points) return;
  const capacity = Math.max(points, scratch.rawFlags.length * 2, 16);
  scratch.rawFlags = new Uint8Array(capacity);
  scratch.encodedFlags = new Uint8Array(capacity);
  scratch.xData = new Uint8Array(capacity * 2);
  scratch.yData = new Uint8Array(capacity * 2);
}

function pointFlags(
  points: DecodedTriplets,
  ctfFlags: Uint8Array,
  scratch: GlyphScratch,
): {
  flagsLength: number;
  xLength: number;
  yLength: number;
} {
  ensureByteScratch(scratch, points.length);
  const rawFlags = scratch.rawFlags;
  const compressedFlags = scratch.encodedFlags;
  const xData = scratch.xData;
  const yData = scratch.yData;
  let xLength = 0;
  let yLength = 0;
  let priorX = 0;
  let priorY = 0;
  for (let i = 0; i < points.length; i++) {
    const px = points.x[i]!;
    const py = points.y[i]!;
    const dx = px - priorX;
    const dy = py - priorY;
    priorX = px;
    priorY = py;
    let flag = (ctfFlags[i]! & 0x80) === 0 ? 0x01 : 0;
    if (dx === 0) flag |= 0x10;
    else if (dx > 0 && dx <= 255) {
      flag |= 0x12;
      xData[xLength++] = dx;
    } else if (dx < 0 && dx >= -255) {
      flag |= 0x02;
      xData[xLength++] = -dx;
    } else {
      xData[xLength++] = dx >>> 8;
      xData[xLength++] = dx;
    }
    if (dy === 0) flag |= 0x20;
    else if (dy > 0 && dy <= 255) {
      flag |= 0x24;
      yData[yLength++] = dy;
    } else if (dy < 0 && dy >= -255) {
      flag |= 0x04;
      yData[yLength++] = -dy;
    } else {
      yData[yLength++] = dy >>> 8;
      yData[yLength++] = dy;
    }
    rawFlags[i] = flag;
  }

  let flagsLength = 0;
  let i = 0;
  while (i < points.length) {
    let run = 1;
    while (i + run < points.length && rawFlags[i + run] === rawFlags[i] && run < 256) run++;
    if (run > 1) {
      compressedFlags[flagsLength++] = rawFlags[i]! | 0x08;
      compressedFlags[flagsLength++] = run - 1;
    } else compressedFlags[flagsLength++] = rawFlags[i]!;
    i += run;
  }
  return { flagsLength, xLength, yLength };
}

function readSimpleGlyph(
  numContours: number,
  explicitBox: [number, number, number, number] | undefined,
  rest: Reader,
  push: Reader,
  code: Reader,
  limits: DecoderLimits,
  glyphId: number,
  writer: Writer,
  scratch: GlyphScratch,
): void {
  // A contourless CTF glyph carries no further data in any stream, and
  // TrueType spells an outline-free glyph as an empty loca range rather than a
  // zero-contour record.
  if (numContours === 0) return;
  if (scratch.endPoints.length < numContours)
    scratch.endPoints = new Uint16Array(Math.max(numContours, scratch.endPoints.length * 2, 8));
  const endPoints = scratch.endPoints;
  let pointCount = 0;
  for (let i = 0; i < numContours; i++) {
    const encoded = read255UShort(rest);
    if (i === 0) pointCount = encoded + 1;
    else pointCount += encoded;
    if (pointCount <= 0 || pointCount > 0x10000)
      fail("INVALID_CTF", "Invalid simple-glyph contour endpoint");
    endPoints[i] = pointCount - 1;
  }
  if (pointCount > limits.maxPointsPerGlyph)
    fail("LIMIT_EXCEEDED", "Glyph point count exceeds configured limit");
  const flags = rest.bytes(pointCount);
  const points = decodeTripletArrays(rest, flags, scratch);
  const [xMin, yMin, xMax, yMax] = explicitBox ?? points.box;
  const encodedPoints = pointFlags(points, flags, scratch);

  writer.i16be(numContours);
  writer.i16be(xMin);
  writer.i16be(yMin);
  writer.i16be(xMax);
  writer.i16be(yMax);
  for (let i = 0; i < numContours; i++) writer.u16be(endPoints[i]!);
  writeInstructions(rest, push, code, limits, glyphId, writer, scratch);
  writer.bytes(scratch.encodedFlags.subarray(0, encodedPoints.flagsLength));
  writer.bytes(scratch.xData.subarray(0, encodedPoints.xLength));
  writer.bytes(scratch.yData.subarray(0, encodedPoints.yLength));
}

function readCompositeGlyph(
  box: [number, number, number, number],
  rest: Reader,
  push: Reader,
  code: Reader,
  limits: DecoderLimits,
  glyphId: number,
  writer: Writer,
  scratch: GlyphScratch,
): void {
  writer.i16be(-1);
  for (const value of box) writer.i16be(value);
  let flags = MORE_COMPONENTS;
  let components = 0;
  while (flags & MORE_COMPONENTS) {
    if (++components > limits.maxGlyphs)
      fail("LIMIT_EXCEEDED", "Composite glyph has too many components");
    flags = rest.u16be();
    writer.u16be(flags);
    writer.u16be(rest.u16be());
    const argsBytes = flags & ARG_1_AND_2_ARE_WORDS ? 4 : 2;
    writer.bytes(rest.bytes(argsBytes));
    if (flags & WE_HAVE_A_TWO_BY_TWO) writer.bytes(rest.bytes(8));
    else if (flags & WE_HAVE_AN_X_AND_Y_SCALE) writer.bytes(rest.bytes(4));
    else if (flags & WE_HAVE_A_SCALE) writer.bytes(rest.bytes(2));
  }
  if (flags & WE_HAVE_INSTRUCTIONS) {
    writeInstructions(rest, push, code, limits, glyphId, writer, scratch);
  }
}

function reconstructGlyphs(
  restData: Uint8Array,
  glyf: DirectoryEntry,
  pushData: Uint8Array,
  codeData: Uint8Array,
  glyphCount: number,
  limits: DecoderLimits,
): { glyf: Uint8Array; offsets: Uint32Array } {
  const rest = new Reader(restData, glyf.offset, glyf.offset + glyf.length);
  const push = new Reader(pushData);
  const code = new Reader(codeData);
  const writer = new Writer(Math.min(glyf.length * 2, limits.maxFontBytes));
  const offsets = new Uint32Array(glyphCount + 1);
  const scratch: GlyphScratch = {
    x: new Int16Array(16),
    y: new Int16Array(16),
    rawFlags: new Uint8Array(16),
    encodedFlags: new Uint8Array(16),
    xData: new Uint8Array(32),
    yData: new Uint8Array(32),
    endPoints: new Uint16Array(8),
    pushValues: new Int16Array(16),
  };

  for (let glyphId = 0; glyphId < glyphCount; glyphId++) {
    offsets[glyphId] = writer.length;
    let numContours = rest.i16be();
    if (numContours < 0) {
      const box: [number, number, number, number] = [
        rest.i16be(),
        rest.i16be(),
        rest.i16be(),
        rest.i16be(),
      ];
      readCompositeGlyph(box, rest, push, code, limits, glyphId, writer, scratch);
    } else {
      let box: [number, number, number, number] | undefined;
      if (numContours === 0x7fff) {
        numContours = rest.i16be();
        if (numContours < 0) fail("INVALID_CTF", "Invalid explicit simple-glyph contour count");
        box = [rest.i16be(), rest.i16be(), rest.i16be(), rest.i16be()];
      }
      readSimpleGlyph(numContours, box, rest, push, code, limits, glyphId, writer, scratch);
    }
    if (writer.length & 1) writer.u8(0);
    if (writer.length > limits.maxFontBytes)
      fail("LIMIT_EXCEEDED", "Reconstructed glyf table exceeds configured limit");
  }
  offsets[glyphCount] = writer.length;
  return { glyf: writer.finish(), offsets };
}

function makeLoca(offsets: Uint32Array, shortFormat: boolean): Uint8Array {
  const writer = new Writer(offsets.length * (shortFormat ? 2 : 4));
  for (const offset of offsets) {
    if (shortFormat) writer.u16be(offset >>> 1);
    else writer.u32be(offset);
  }
  return writer.finish();
}

export function parseCtf(
  streams: readonly [Uint8Array, Uint8Array, Uint8Array],
  limits: DecoderLimits,
  onWarn?: (message: string) => void,
): SfntContainer {
  const restData = streams[0];
  const directory = new Reader(restData);
  const sfntVersion = directory.u32be();
  const tableCount = directory.u16be();
  if (tableCount === 0 || tableCount > limits.maxTables)
    fail("LIMIT_EXCEEDED", `Invalid CTF table count ${tableCount}`);
  directory.skip(6);
  const entries: DirectoryEntry[] = [];
  const byTag = new Map<string, DirectoryEntry>();
  for (let i = 0; i < tableCount; i++) {
    const tag = tagAt(restData, directory.pos);
    // Tag lookups below take the last entry while the emit loop keeps every
    // entry, so a duplicate would resolve head/glyf to one table yet write two
    // directory records for it. validateSfnt does not check tag uniqueness.
    if (byTag.has(tag)) fail("INVALID_CTF", `Duplicate CTF table tag ${tag}`);
    directory.skip(4);
    directory.u32be();
    const offset = directory.u32be();
    const length = directory.u32be();
    const entry = { tag, offset, length };
    entries.push(entry);
    byTag.set(tag, entry);
  }

  const headEntry = byTag.get("head");
  const maxpEntry = byTag.get("maxp");
  const glyfEntry = byTag.get("glyf");
  const locaEntry = byTag.get("loca");
  if (!headEntry || !maxpEntry || !glyfEntry || !locaEntry)
    fail("INVALID_CTF", "CTF is missing head, maxp, glyf, or loca");
  // head is the only table this rewrites, so it is the only one that needs its
  // own storage; the rest stay views into the decompressed stream until
  // buildSfnt copies them into the font.
  const head = tableView(restData, headEntry).slice();
  const maxp = tableView(restData, maxpEntry);
  if (head.length < 54 || maxp.length < 6) fail("INVALID_CTF", "Malformed head or maxp table");
  head.fill(0, 8, 12);
  const glyphCount = new DataView(maxp.buffer, maxp.byteOffset, maxp.byteLength).getUint16(
    4,
    false,
  );
  if (glyphCount > limits.maxGlyphs) fail("LIMIT_EXCEEDED", `Font declares ${glyphCount} glyphs`);

  const reconstructed = reconstructGlyphs(
    restData,
    glyfEntry,
    streams[1],
    streams[2],
    glyphCount,
    limits,
  );
  const oldLocaFormat = new DataView(head.buffer, head.byteOffset, head.byteLength).getInt16(
    50,
    false,
  );
  const useShortLoca = oldLocaFormat === 0 && reconstructed.glyf.length / 2 <= 0xffff;
  new DataView(head.buffer, head.byteOffset, head.byteLength).setInt16(
    50,
    useShortLoca ? 0 : 1,
    false,
  );
  const loca = makeLoca(reconstructed.offsets, useShortLoca);

  const droppedTables: string[] = [];
  const tables: SfntTable[] = [];
  for (const entry of entries) {
    let data: Uint8Array;
    if (entry.tag === "glyf") data = reconstructed.glyf;
    else if (entry.tag === "loca") data = loca;
    else if (entry.tag === "head") data = head;
    else if (entry.tag === "cvt ") data = decodeCvt(tableView(restData, entry));
    else if (entry.tag === "hdmx" || entry.tag === "VDMX") {
      droppedTables.push(entry.tag);
      onWarn?.(`Dropped optional ${entry.tag} table; outlines and metrics remain usable.`);
      continue;
    } else data = tableView(restData, entry);
    tables.push({ tag: entry.tag, data });
  }
  return { sfntVersion, tables, ...(droppedTables.length ? { droppedTables } : {}) };
}
