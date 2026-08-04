/**
 * Test-only MTX/EOT encoder.
 *
 * The decoder in `src/fonts/mtx` has no fixtures to work against: every
 * embedded font we ship is an uncompressed EOT, so LZCOMP and CTF are
 * otherwise dead code. This module builds compressed fixtures from the
 * W3C MicroType Express specification so those layers can be round-tripped.
 *
 * It is deliberately naive — it optimizes nothing and only needs to emit
 * streams a conforming decoder must accept.
 */

import { AdaptiveHuffman } from "../../fonts/mtx/lzcomp";
import { tripletEncoding } from "../../fonts/mtx/triplet";

// ── Bit output ──────────────────────────────────────────────────────

export class BitWriter {
  private readonly bits: number[] = [];

  bit(value: number): void {
    this.bits.push(value & 1);
  }

  value(value: number, width: number): void {
    for (let i = width - 1; i >= 0; i--) this.bit(value / 2 ** i);
  }

  codes(values: readonly number[]): void {
    for (const value of values) this.bits.push(value & 1);
  }

  finish(): Uint8Array {
    const out = new Uint8Array(Math.ceil(this.bits.length / 8));
    for (let i = 0; i < this.bits.length; i++) {
      out[i >>> 3]! |= this.bits[i]! << (7 - (i & 7));
    }
    return out;
  }
}

// ── LZCOMP ──────────────────────────────────────────────────────────

/** A back-reference, expressed the way the decoder reconstructs it. */
export interface CopyOp {
  /** Distance from the tail of the copied phrase, as encoded (>= 1). */
  distance: number;
  length: number;
}

export type LzcompOp = { literal: number } | { dup: 2 | 4 | 6 } | CopyOp;

function distanceRangesFor(length: number): number {
  let ranges = 1;
  while (2 ** (3 * ranges) < length) ranges++;
  return ranges;
}

/**
 * Encode an explicit op list. `outputLength` is what the decoder will be told
 * to produce, which determines the symbol alphabet, so callers must pass the
 * length the ops actually expand to.
 */
export interface LzcompOptions {
  runLength?: boolean;
  /** Version 1 streams omit the leading run-length flag bit. */
  version?: 1 | 3;
}

export function encodeLzcompOps(
  ops: readonly LzcompOp[],
  outputLength: number,
  options: LzcompOptions = {},
): Uint8Array {
  const ranges = distanceRangesFor(outputLength);
  const dup2 = 256 + 8 * ranges;
  const distanceTree = new AdaptiveHuffman(8);
  const lengthTree = new AdaptiveHuffman(8);
  const symbolTree = new AdaptiveHuffman(dup2 + 3);

  const bits = new BitWriter();
  if ((options.version ?? 3) !== 1) bits.bit(options.runLength ? 1 : 0);
  bits.value(outputLength, 24);

  for (const op of ops) {
    if ("literal" in op) {
      bits.codes(symbolTree.encode(op.literal));
      continue;
    }
    if ("dup" in op) {
      bits.codes(symbolTree.encode(dup2 + (op.dup / 2 - 1)));
      continue;
    }

    // The decoder adds one to the length of any far copy before using it.
    const encodedLength = op.length - 2 - (op.distance >= 512 ? 1 : 0);
    if (encodedLength < 0)
      throw new Error(`Copy length ${op.length} is too short for distance ${op.distance}`);

    // Length travels as base-4 digits, most significant first: the top digit
    // rides in the symbol's low bits, the rest come from the length tree.
    const digits: number[] = [];
    let remainder = encodedLength;
    do {
      digits.unshift(remainder & 3);
      remainder >>>= 2;
    } while (remainder > 0);

    let used = 1;
    while (op.distance - 1 >= 8 ** used) used++;
    if (used > ranges)
      throw new Error(`Distance ${op.distance} needs more ranges than the stream declares`);

    const head = (digits.length > 1 ? 4 : 0) | digits[0]!;
    bits.codes(symbolTree.encode(256 + (used - 1) * 8 + head));
    for (let i = 1; i < digits.length; i++) {
      bits.codes(lengthTree.encode((i < digits.length - 1 ? 4 : 0) | digits[i]!));
    }

    for (let i = used - 1; i >= 0; i--) {
      bits.codes(distanceTree.encode(Math.floor((op.distance - 1) / 8 ** i) % 8));
    }
  }

  return bits.finish();
}

/** Literal-only encoding: the simplest stream a decoder must accept. */
export function encodeLzcompLiterals(data: Uint8Array, options: LzcompOptions = {}): Uint8Array {
  return encodeLzcompOps(
    Array.from(data, (literal) => ({ literal })),
    data.length,
    options,
  );
}

/** Bound on the greedy match search, which is otherwise quadratic. */
const SEARCH_WINDOW = 512;

/**
 * Greedy LZ77 over the same history the decoder maintains, so the copy path
 * gets exercised with real back-references rather than hand-written ops.
 */
export function encodeLzcompGreedy(data: Uint8Array, options: LzcompOptions = {}): Uint8Array {
  const ranges = distanceRangesFor(data.length);
  const maxDistance = Math.min(8 ** ranges, SEARCH_WINDOW);
  const ops: LzcompOp[] = [];

  for (let pos = 0; pos < data.length;) {
    let bestLength = 0;
    let bestDistance = 0;
    const earliest = Math.max(0, pos - maxDistance);
    for (let start = earliest; start < pos; start++) {
      let length = 0;
      while (
        length < pos - start &&
        pos + length < data.length &&
        data[start + length] === data[pos + length]
      )
        length++;
      // A copy's source must end at or before the current position, which the
      // `length <= pos - start` bound above already guarantees.
      const distance = pos - start - length + 1;
      if (distance < 1 || distance > maxDistance) continue;
      const minimum = distance >= 512 ? 3 : 2;
      if (length >= minimum && length > bestLength) {
        bestLength = length;
        bestDistance = distance;
      }
    }

    if (bestLength >= 3) {
      ops.push({ distance: bestDistance, length: bestLength });
      pos += bestLength;
    } else {
      ops.push({ literal: data[pos]! });
      pos++;
    }
  }

  return encodeLzcompOps(ops, data.length, options);
}

/** Pack bytes with the optional run-length layer the decoder can expand. */
export function packRunLength(data: Uint8Array, escape: number): Uint8Array {
  const out: number[] = [escape];
  for (let i = 0; i < data.length;) {
    const value = data[i]!;
    let run = 1;
    while (run < 255 && i + run < data.length && data[i + run] === value) run++;
    if (run >= 3 && value !== escape) {
      out.push(escape, run, value);
    } else {
      for (let j = 0; j < run; j++) {
        if (value === escape) out.push(escape, 0);
        else out.push(value);
      }
    }
    i += run;
  }
  return Uint8Array.from(out);
}

// ── MTX container ───────────────────────────────────────────────────

export function buildMtx(
  streams: readonly [Uint8Array, Uint8Array, Uint8Array],
  version = 3,
): Uint8Array {
  const offset2 = 10 + streams[0].length;
  const offset3 = offset2 + streams[1].length;
  const total = offset3 + streams[2].length;
  const out = new Uint8Array(total);
  const view = new DataView(out.buffer);
  out[0] = version;
  const u24 = (at: number, value: number): void => {
    out[at] = value >>> 16;
    view.setUint16(at + 1, value & 0xffff, false);
  };
  u24(1, 10);
  u24(4, offset2);
  u24(7, offset3);
  out.set(streams[0], 10);
  out.set(streams[1], offset2);
  out.set(streams[2], offset3);
  return out;
}

// ── EOT container ───────────────────────────────────────────────────

export const TTEMBED_TTCOMPRESSED = 0x00000004;
export const TTEMBED_XORENCRYPTDATA = 0x10000000;

/** Wrap font data in a version-1 EOT with four empty name strings. */
export function buildEot(fontData: Uint8Array, flags: number): Uint8Array {
  const headerSize = 96;
  const payload =
    flags & TTEMBED_XORENCRYPTDATA ? Uint8Array.from(fontData, (byte) => byte ^ 0x50) : fontData;
  const out = new Uint8Array(headerSize + payload.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, out.length, true);
  view.setUint32(4, payload.length, true);
  view.setUint32(8, 0x00010000, true);
  view.setUint32(12, flags, true);
  view.setUint32(28, 400, true); // Weight
  view.setUint16(32, 0, true); // fsType
  view.setUint16(34, 0x504c, true); // MagicNumber
  out.set(payload, headerSize);
  return out;
}

// ── CTF value encodings ─────────────────────────────────────────────

export function write255UShort(out: number[], value: number): void {
  if (value < 253) out.push(value);
  else if (value <= 508) out.push(255, value - 253);
  else if (value <= 761) out.push(254, value - 506);
  else out.push(253, (value >>> 8) & 0xff, value & 0xff);
}

export function write255Short(out: number[], value: number): void {
  const magnitude = Math.abs(value);
  // Code 253 carries its own sign, and the decoder rejects it after a 250
  // prefix, so wide values must skip the prefix entirely.
  if (magnitude > 755) {
    out.push(253, (value >> 8) & 0xff, value & 0xff);
    return;
  }
  if (value < 0) out.push(250);
  if (magnitude <= 249) out.push(magnitude);
  else if (magnitude <= 505) out.push(255, magnitude - 250);
  else out.push(254, magnitude - 500);
}

export function encodeCvt(values: readonly number[]): Uint8Array {
  const out: number[] = [(values.length >>> 8) & 0xff, values.length & 0xff];
  let previous = 0;
  for (const value of values) {
    const delta = (((value - previous) & 0xffff) ^ 0x8000) - 0x8000;
    previous = value & 0xffff;
    if (delta >= 0 && delta <= 237) {
      out.push(delta);
      continue;
    }
    let done = false;
    for (let j = 1; j <= 8 && !done; j++) {
      const rest = delta - 238 * j;
      if (rest >= 0 && rest <= 255) {
        out.push(247 + j, rest);
        done = true;
      }
    }
    for (let j = 0; j <= 8 && !done; j++) {
      const rest = -delta - 238 * j;
      if (rest >= 0 && rest <= 255) {
        out.push(239 + j, rest);
        done = true;
      }
    }
    if (!done) out.push(238, (delta >> 8) & 0xff, delta & 0xff);
  }
  return Uint8Array.from(out);
}

/** Pick the most compact triplet index that represents (dx, dy) exactly. */
export function encodeTriplet(dx: number, dy: number, onCurve: boolean): number[] {
  for (let index = 0; index < 128; index++) {
    const encoding = tripletEncoding(index);
    const rawX = (encoding.xNegative ? -dx : dx) - encoding.deltaX;
    const rawY = (encoding.yNegative ? -dy : dy) - encoding.deltaY;
    if (rawX < 0 || rawX >= 2 ** encoding.xBits) continue;
    if (rawY < 0 || rawY >= 2 ** encoding.yBits) continue;

    const out = [index | (onCurve ? 0 : 0x80)];
    const packed = rawX * 2 ** encoding.yBits + rawY;
    for (let i = encoding.bytes - 2; i >= 0; i--) {
      out.push(Math.floor(packed / 2 ** (8 * i)) % 256);
    }
    return out;
  }
  throw new Error(`No triplet encoding represents (${dx}, ${dy})`);
}

// ── CTF font model ──────────────────────────────────────────────────

export interface GlyphPoint {
  x: number;
  y: number;
  onCurve: boolean;
}

export interface GlyphComponent {
  flags: number;
  glyphIndex: number;
  /** Raw argument bytes, matching the width implied by `flags`. */
  args: number[];
}

export interface GlyphHints {
  /** Values reconstructed into a PUSH burst ahead of the code stream. */
  pushes?: number[];
  /** Raw bytecode appended after the pushes. */
  code?: number[];
  /** Emit hop codes instead of plain values where the pattern allows. */
  useHopCodes?: boolean;
}

export interface SimpleGlyph extends GlyphHints {
  kind: "simple";
  contours: GlyphPoint[][];
  /** Forces the explicit-bounding-box form (numContours == 0x7FFF). */
  bbox?: [number, number, number, number];
}

export interface CompositeGlyph extends GlyphHints {
  kind: "composite";
  bbox: [number, number, number, number];
  components: GlyphComponent[];
}

export type CtfGlyph = SimpleGlyph | CompositeGlyph;

export interface CtfFontSpec {
  glyphs: CtfGlyph[];
  cvt?: number[];
  /** Extra tables copied through stream 0 untouched. */
  extraTables?: { tag: string; data: Uint8Array }[];
  indexToLocFormat?: 0 | 1;
  unitsPerEm?: number;
}

function pushInt16(out: number[], value: number): void {
  out.push((value >> 8) & 0xff, value & 0xff);
}

function encodeHints(glyph: GlyphHints, rest: number[], push: number[], code: number[]): void {
  const values = glyph.pushes ?? [];
  const bytecode = glyph.code ?? [];
  write255UShort(rest, values.length);
  write255UShort(rest, bytecode.length);

  for (let i = 0; i < values.length;) {
    // Hop codes replay the value two slots back, so they need that history and
    // enough room left in the declared count.
    const canHop = glyph.useHopCodes === true && i >= 2 && values[i] === values[i - 2];
    if (
      canHop &&
      i + 5 <= values.length &&
      values[i + 2] === values[i] &&
      values[i + 4] === values[i]
    ) {
      push.push(252);
      write255Short(push, values[i + 1]!);
      write255Short(push, values[i + 3]!);
      i += 5;
    } else if (canHop && i + 3 <= values.length && values[i + 2] === values[i]) {
      push.push(251);
      write255Short(push, values[i + 1]!);
      i += 3;
    } else {
      write255Short(push, values[i]!);
      i++;
    }
  }
  code.push(...bytecode);
}

function encodeGlyph(glyph: CtfGlyph, rest: number[], push: number[], code: number[]): void {
  if (glyph.kind === "composite") {
    pushInt16(rest, -1);
    for (const value of glyph.bbox) pushInt16(rest, value);
    glyph.components.forEach((component, index) => {
      const last = index === glyph.components.length - 1;
      const flags = last ? component.flags & ~0x0020 : component.flags | 0x0020;
      pushInt16(rest, flags);
      pushInt16(rest, component.glyphIndex);
      rest.push(...component.args);
    });
    const trailing = glyph.components.at(-1)?.flags ?? 0;
    if (trailing & 0x0100) encodeHints(glyph, rest, push, code);
    return;
  }

  if (glyph.contours.length === 0) {
    pushInt16(rest, 0);
    return;
  }

  if (glyph.bbox) {
    pushInt16(rest, 0x7fff);
    pushInt16(rest, glyph.contours.length);
    for (const value of glyph.bbox) pushInt16(rest, value);
  } else {
    pushInt16(rest, glyph.contours.length);
  }

  glyph.contours.forEach((contour, index) => {
    if (contour.length === 0) throw new Error("A CTF contour must contain at least one point");
    write255UShort(rest, index === 0 ? contour.length - 1 : contour.length);
  });

  const points = glyph.contours.flat();
  const coordinates: number[] = [];
  let priorX = 0;
  let priorY = 0;
  for (const point of points) {
    const triplet = encodeTriplet(point.x - priorX, point.y - priorY, point.onCurve);
    priorX = point.x;
    priorY = point.y;
    rest.push(triplet[0]!);
    coordinates.push(...triplet.slice(1));
  }
  rest.push(...coordinates);
  encodeHints(glyph, rest, push, code);
}

function makeHead(indexToLocFormat: number, unitsPerEm: number): Uint8Array {
  const head = new Uint8Array(54);
  const view = new DataView(head.buffer);
  view.setUint32(0, 0x00010000, false);
  view.setUint32(4, 0x00010000, false); // fontRevision
  view.setUint32(8, 0xdeadbeef, false); // checksumAdjustment, must be zeroed
  view.setUint32(12, 0x5f0f3cf5, false); // magicNumber
  view.setUint16(16, 0x000b, false); // flags
  view.setUint16(18, unitsPerEm, false);
  view.setInt16(50, indexToLocFormat, false);
  return head;
}

function makeMaxp(numGlyphs: number): Uint8Array {
  const maxp = new Uint8Array(32);
  const view = new DataView(maxp.buffer);
  view.setUint32(0, 0x00010000, false);
  view.setUint16(4, numGlyphs, false);
  return maxp;
}

/** Build the three CTF streams for a font description. */
export function buildCtfStreams(spec: CtfFontSpec): [Uint8Array, Uint8Array, Uint8Array] {
  const glyphBytes: number[] = [];
  const push: number[] = [];
  const code: number[] = [];
  for (const glyph of spec.glyphs) encodeGlyph(glyph, glyphBytes, push, code);

  const tables: { tag: string; data: Uint8Array }[] = [
    { tag: "head", data: makeHead(spec.indexToLocFormat ?? 0, spec.unitsPerEm ?? 1000) },
    { tag: "maxp", data: makeMaxp(spec.glyphs.length) },
    ...(spec.cvt ? [{ tag: "cvt ", data: encodeCvt(spec.cvt) }] : []),
    ...(spec.extraTables ?? []),
    { tag: "glyf", data: Uint8Array.from(glyphBytes) },
    { tag: "loca", data: new Uint8Array(0) },
  ];

  const directorySize = 12 + tables.length * 16;
  let dataSize = 0;
  for (const table of tables) dataSize += table.data.length;

  const stream = new Uint8Array(directorySize + dataSize);
  const view = new DataView(stream.buffer);
  view.setUint32(0, 0x00010000, false);
  view.setUint16(4, tables.length, false);

  let offset = directorySize;
  tables.forEach((table, index) => {
    const record = 12 + index * 16;
    for (let i = 0; i < 4; i++) stream[record + i] = table.tag.charCodeAt(i);
    view.setUint32(record + 4, 0, false);
    view.setUint32(record + 8, table.tag === "loca" ? 0 : offset, false);
    view.setUint32(record + 12, table.data.length, false);
    stream.set(table.data, offset);
    offset += table.data.length;
  });

  return [stream, Uint8Array.from(push), Uint8Array.from(code)];
}

export interface BuildEotOptions extends LzcompOptions {
  encrypted?: boolean;
  /** Skip match finding; useful for large fixtures where the search dominates. */
  literalOnly?: boolean;
}

/** Full pipeline: font description to a compressed, EOT-wrapped fixture. */
export function buildCompressedEot(spec: CtfFontSpec, options: BuildEotOptions = {}): Uint8Array {
  const encode = options.literalOnly ? encodeLzcompLiterals : encodeLzcompGreedy;
  const streams = buildCtfStreams(spec).map((stream) => encode(stream, options));
  const flags = TTEMBED_TTCOMPRESSED | (options.encrypted ? TTEMBED_XORENCRYPTDATA : 0);
  return buildEot(
    buildMtx(streams as [Uint8Array, Uint8Array, Uint8Array], options.version),
    flags,
  );
}
