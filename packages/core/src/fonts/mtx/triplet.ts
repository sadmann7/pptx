import { Reader } from "./binary";
import { fail } from "./error";

export interface GlyphPoint {
  x: number;
  y: number;
  onCurve: boolean;
}

interface Encoding {
  bytes: number;
  xBits: number;
  yBits: number;
  deltaX: number;
  deltaY: number;
  xNegative: boolean;
  yNegative: boolean;
}

const BYTE_COUNT = new Uint8Array(128);
const X_SHIFT = new Uint8Array(128);
const Y_SHIFT = new Uint8Array(128);
const X_MASK = new Uint16Array(128);
const Y_MASK = new Uint16Array(128);
const DELTA_X = new Uint16Array(128);
const DELTA_Y = new Uint16Array(128);
const SIGNS = new Uint8Array(128); // bit 0 = negative x, bit 1 = negative y

function setEncoding(index: number, encoding: Encoding): void {
  BYTE_COUNT[index] = encoding.bytes;
  const totalBits = (encoding.bytes - 1) * 8;
  X_SHIFT[index] = totalBits - encoding.xBits;
  Y_SHIFT[index] = totalBits - encoding.xBits - encoding.yBits;
  X_MASK[index] = encoding.xBits === 0 ? 0 : 2 ** encoding.xBits - 1;
  Y_MASK[index] = encoding.yBits === 0 ? 0 : 2 ** encoding.yBits - 1;
  DELTA_X[index] = encoding.deltaX;
  DELTA_Y[index] = encoding.deltaY;
  SIGNS[index] = (encoding.xNegative ? 1 : 0) | (encoding.yNegative ? 2 : 0);
}

export function tripletEncoding(index: number): Encoding {
  if (index < 0 || index > 127) fail("INVALID_CTF", `Invalid triplet index ${index}`);
  if (index < 10)
    return {
      bytes: 2,
      xBits: 0,
      yBits: 8,
      deltaX: 0,
      deltaY: (index >>> 1) * 256,
      xNegative: false,
      yNegative: (index & 1) === 0,
    };
  if (index < 20) {
    const local = index - 10;
    return {
      bytes: 2,
      xBits: 8,
      yBits: 0,
      deltaX: (local >>> 1) * 256,
      deltaY: 0,
      xNegative: (local & 1) === 0,
      yNegative: false,
    };
  }
  if (index < 84) {
    const local = index - 20;
    const signs = local & 3;
    return {
      bytes: 2,
      xBits: 4,
      yBits: 4,
      deltaX: 1 + 16 * Math.floor(local / 16),
      deltaY: 1 + 16 * Math.floor((local % 16) / 4),
      xNegative: (signs & 1) === 0,
      yNegative: (signs & 2) === 0,
    };
  }
  if (index < 120) {
    const local = index - 84;
    const signs = local & 3;
    return {
      bytes: 3,
      xBits: 8,
      yBits: 8,
      deltaX: 1 + 256 * Math.floor(local / 12),
      deltaY: 1 + 256 * Math.floor((local % 12) / 4),
      xNegative: (signs & 1) === 0,
      yNegative: (signs & 2) === 0,
    };
  }
  const signs = index & 3;
  return {
    bytes: index < 124 ? 4 : 5,
    xBits: index < 124 ? 12 : 16,
    yBits: index < 124 ? 12 : 16,
    deltaX: 0,
    deltaY: 0,
    xNegative: (signs & 1) === 0,
    yNegative: (signs & 2) === 0,
  };
}

for (let i = 0; i < 128; i++) setEncoding(i, tripletEncoding(i));

export interface DecodedTriplets {
  x: Int16Array;
  y: Int16Array;
  length: number;
  box: [number, number, number, number];
}

export interface TripletScratch {
  x: Int16Array;
  y: Int16Array;
}

export function decodeTripletArrays(
  reader: Reader,
  flags: Uint8Array,
  scratch?: TripletScratch,
): DecodedTriplets {
  const target = scratch ?? { x: new Int16Array(flags.length), y: new Int16Array(flags.length) };
  if (target.x.length < flags.length) {
    const capacity = Math.max(flags.length, target.x.length * 2, 16);
    target.x = new Int16Array(capacity);
    target.y = new Int16Array(capacity);
  }
  const xs = target.x;
  const ys = target.y;
  let x = 0;
  let y = 0;
  let xMin = 0;
  let yMin = 0;
  let xMax = 0;
  let yMax = 0;
  // Bounds are checked once per point against a local cursor rather than once
  // per byte through the reader.
  const data = reader.data;
  const end = reader.end;
  let pos = reader.pos;
  for (let i = 0; i < flags.length; i++) {
    const flag = flags[i]!;
    const index = flag & 0x7f;
    const byteCount = BYTE_COUNT[index]!;
    if (pos + byteCount - 1 > end) fail("BOUNDS", "Unexpected end of triplet data", pos);
    let packed = 0;
    for (let j = 1; j < byteCount; j++) packed = packed * 256 + data[pos++]!;
    let dx = ((packed >>> X_SHIFT[index]!) & X_MASK[index]!) + DELTA_X[index]!;
    let dy = ((packed >>> Y_SHIFT[index]!) & Y_MASK[index]!) + DELTA_Y[index]!;
    const signs = SIGNS[index]!;
    if (signs & 1) dx = -dx;
    if (signs & 2) dy = -dy;
    x += dx;
    y += dy;
    if (x < -32768 || x > 32767 || y < -32768 || y > 32767) {
      fail("INVALID_CTF", "Glyph coordinate exceeds TrueType range", pos);
    }
    xs[i] = x;
    ys[i] = y;
    if (i === 0) {
      xMin = x;
      xMax = x;
      yMin = y;
      yMax = y;
    } else {
      if (x < xMin) xMin = x;
      if (x > xMax) xMax = x;
      if (y < yMin) yMin = y;
      if (y > yMax) yMax = y;
    }
  }
  reader.pos = pos;
  return { x: xs, y: ys, length: flags.length, box: [xMin, yMin, xMax, yMax] };
}
