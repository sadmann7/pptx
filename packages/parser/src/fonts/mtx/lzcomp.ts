/*
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * Derived from mtx-decompressor v1.4.2
 * (https://github.com/ChristopherVR/mtx-decompressor, © ChristopherVR,
 * MPL-2.0). Optimized for @diceui/pptx-parser:
 *  - adaptive Huffman trees as flat typed arrays (no per-node objects)
 *  - bit reader inlined into the decoder without per-bit error dispatch
 *  - non-RLE output taken directly from the LZ window (the original wrote
 *    every byte a second time through a per-byte closure)
 * Output is byte-identical to the original.
 */

const LEN_WIDTH = 3;
const DIST_WIDTH = 3;
const BIT_RANGE = LEN_WIDTH - 1;
const MAX_2BYTE_DIST = 512;
const PRELOAD_SIZE = 2 * 32 * 96 + 4 * 256;
const LEN_MIN = 2;
const DIST_MIN = 1;
const MAX_OUT_LEN = 4 * 1024 * 1024;
const MAX_OUT = 16 * 1024 * 1024;

// ── Bit input ───────────────────────────────────────────────────────

/**
 * MSB-first bit reader over a byte range with a 24-bit prefetch buffer.
 *
 * Fields are intentionally public: the adaptive-Huffman decoder walks its
 * tree with the buffer held in locals to avoid per-bit call overhead.
 * Prefetching whole bytes early is safe — like the byte-at-a-time original,
 * it only fails when a bit is actually requested past the end of data.
 */
class BitReader {
  readonly data: Uint8Array;
  index: number;
  readonly size: number;
  /** Low `bitCount` bits are valid, MSB-first (next bit is the highest). */
  bitBuffer = 0;
  bitCount = 0;

  constructor(data: Uint8Array, offset: number, size: number) {
    this.data = data;
    this.index = offset;
    this.size = size;
  }

  /** Load bytes until at least one bit is available (throws when exhausted). */
  refill(): void {
    if (this.index >= this.size) {
      throw new Error("BitIO: end of data");
    }
    do {
      this.bitBuffer = ((this.bitBuffer << 8) | this.data[this.index++]) >>> 0;
      this.bitCount += 8;
    } while (this.bitCount <= 16 && this.index < this.size);
  }

  /** Read a single bit (MSB first). Returns 0 or 1. */
  inputBit(): number {
    if (this.bitCount === 0) this.refill();
    return (this.bitBuffer >>> --this.bitCount) & 1;
  }

  /** Read an unsigned integer of `numberOfBits` width, MSB first. */
  readValue(numberOfBits: number): number {
    let value = 0;
    for (let i = 0; i < numberOfBits; i++) {
      value = ((value << 1) | this.inputBit()) >>> 0;
    }
    return value;
  }
}

// ── Adaptive Huffman (flat typed-array layout) ──────────────────────

const ROOT = 1;

/**
 * Adaptive Huffman decoder over a BitReader.
 *
 * The node tree is stored as parallel Int32Arrays indexed by node id
 * (structure-of-arrays): cache-friendly and free of per-node allocations.
 * Semantics mirror the reference MTX_AHUFF implementation exactly, including
 * the initial weight seeding, so bitstream positions stay in lockstep with
 * the original decoder.
 */
class AHuff {
  private readonly bio: BitReader;
  private readonly up: Int32Array;
  private readonly left: Int32Array;
  private readonly right: Int32Array;
  private readonly code: Int32Array;
  private readonly weight: Int32Array;
  private readonly symbolIndex: Int32Array;

  constructor(bio: BitReader, range: number) {
    this.bio = bio;

    const treeSize = 2 * range;
    const up = new Int32Array(treeSize);
    const left = new Int32Array(treeSize);
    const right = new Int32Array(treeSize);
    const code = new Int32Array(treeSize).fill(-1);
    const weight = new Int32Array(treeSize);

    for (let i = 2; i < treeSize; i++) {
      up[i] = i >> 1;
      weight[i] = 1;
    }
    for (let i = 1; i < range; i++) {
      left[i] = 2 * i;
      right[i] = 2 * i + 1;
    }
    const symbolIndex = new Int32Array(range);
    for (let i = 0; i < range; i++) {
      const leafIdx = range + i;
      code[leafIdx] = i;
      left[leafIdx] = -1;
      right[leafIdx] = -1;
      symbolIndex[i] = leafIdx;
    }

    this.up = up;
    this.left = left;
    this.right = right;
    this.code = code;
    this.weight = weight;
    this.symbolIndex = symbolIndex;

    this.initWeight(ROOT);

    // Initial weight seeding (must match the reference implementation).
    const bitCount2 = range > 256 && range < 512 ? 1 : 0;
    if (bitCount2 !== 0) {
      this.updateWeight(symbolIndex[256]);
      this.updateWeight(symbolIndex[257]);
      const dup2Sym = range - 3;
      for (let i = 0; i < 12; i++) this.updateWeight(symbolIndex[dup2Sym]);
      const dup4Sym = range - 2;
      for (let i = 0; i < 6; i++) this.updateWeight(symbolIndex[dup4Sym]);
    } else {
      for (let j = 0; j < 2; j++) {
        for (let i = 0; i < range; i++) this.updateWeight(symbolIndex[i]);
      }
    }
  }

  /** Decode one symbol: walk the tree bit by bit, then update weights. */
  readSymbol(): number {
    const { left, right, code, bio } = this;
    // Hold the bit buffer in locals for the descent (hot path).
    let bb = bio.bitBuffer;
    let bc = bio.bitCount;
    let a = ROOT;
    let symbol: number;
    do {
      if (bc === 0) {
        bio.bitBuffer = bb;
        bio.bitCount = bc;
        bio.refill();
        bb = bio.bitBuffer;
        bc = bio.bitCount;
      }
      a = (bb >>> --bc) & 1 ? right[a] : left[a];
      symbol = code[a];
    } while (symbol < 0);
    bio.bitBuffer = bb;
    bio.bitCount = bc;
    this.updateWeight(a);
    return symbol;
  }

  /**
   * Increment the weight of node `a` and propagate up to ROOT, swapping
   * nodes to maintain the sibling property (non-increasing weight by index).
   */
  private updateWeight(a: number): void {
    const { up, weight } = this;
    for (; a !== ROOT; a = up[a]) {
      const weightA = weight[a];
      let b = a - 1;
      if (weight[b] === weightA) {
        do {
          b--;
        } while (weight[b] === weightA);
        b++;
        if (b > ROOT) {
          this.swapNodes(a, b);
          a = b;
        }
      }
      weight[a] = weightA + 1;
    }
    weight[ROOT]++;
  }

  /** Swap node contents; `up` pointers stay with the positions. */
  private swapNodes(a: number, b: number): void {
    const { up, left, right, code, weight, symbolIndex } = this;

    let tmp = left[a];
    left[a] = left[b];
    left[b] = tmp;
    tmp = right[a];
    right[a] = right[b];
    right[b] = tmp;
    tmp = code[a];
    code[a] = code[b];
    code[b] = tmp;
    tmp = weight[a];
    weight[a] = weight[b];
    weight[b] = tmp;

    let c = code[a];
    if (c < 0) {
      up[left[a]] = a;
      up[right[a]] = a;
    } else {
      symbolIndex[c] = a;
    }
    c = code[b];
    if (c < 0) {
      up[left[b]] = b;
      up[right[b]] = b;
    } else {
      symbolIndex[c] = b;
    }
  }

  private initWeight(a: number): number {
    if (this.code[a] >= 0) return this.weight[a];
    const w = this.initWeight(this.left[a]) + this.initWeight(this.right[a]);
    this.weight[a] = w;
    return w;
  }
}

// ── LZCOMP ──────────────────────────────────────────────────────────

function setDistRange(length: number): {
  DUP2: number;
  DUP4: number;
  DUP6: number;
  NUM_SYMS: number;
} {
  let numDistRanges = 1;
  let distMax = DIST_MIN + ((1 << (DIST_WIDTH * numDistRanges)) - 1);
  while (distMax < length) {
    numDistRanges++;
    if (numDistRanges > 8) {
      throw new Error("LZCOMP setDistRange: numDistRanges exceeds bound (8)");
    }
    distMax = DIST_MIN + ((1 << (DIST_WIDTH * numDistRanges)) - 1);
  }
  const DUP2 = 256 + (1 << LEN_WIDTH) * numDistRanges;
  return { DUP2, DUP4: DUP2 + 1, DUP6: DUP2 + 2, NUM_SYMS: DUP2 + 3 };
}

function initializeModel(window: Uint8Array): void {
  let i = 0;
  for (let k = 0; k < 32; k++) {
    for (let j = 0; j < 96; j++) {
      window[i++] = k;
      window[i++] = j;
    }
  }
  let j = 0;
  while (i < PRELOAD_SIZE && j < 256) {
    window[i++] = j;
    window[i++] = j;
    window[i++] = j;
    window[i++] = j;
    j++;
  }
}

/** Side channel for decodeLength (single-threaded; avoids a per-match object). */
let lastNumDistRanges = 0;

function decodeLength(lenEcoder: AHuff, symbol: number): number {
  const mask = 1 << BIT_RANGE;
  let bits = symbol - 256;
  lastNumDistRanges = ((bits / (1 << LEN_WIDTH)) | 0) + 1;
  bits %= 1 << LEN_WIDTH;

  let value = 0;
  let iters = 0;
  for (;;) {
    if (++iters > 16) {
      throw new Error("LZCOMP decodeLength: iteration cap exceeded");
    }
    const done = (bits & mask) === 0;
    value = (value << BIT_RANGE) | (bits & ~mask);
    if (done) break;
    bits = lenEcoder.readSymbol();
  }
  return value + LEN_MIN;
}

function decodeDistance(distEcoder: AHuff, numDistRanges: number): number {
  let value = 0;
  for (let i = numDistRanges; i > 0; i--) {
    value = (value << DIST_WIDTH) | distEcoder.readSymbol();
  }
  return value + DIST_MIN;
}

/** RLE post-pass decoder used when the stream declares run-length packing. */
class RleSink {
  private out: Uint8Array;
  private outSize: number;
  private outIdx = 0;
  private state = 0; // 0 initial, 1 normal, 2 seen-escape, 3 need-byte
  private escape = 0;
  private count = 0;

  constructor(capacity: number) {
    this.out = new Uint8Array(capacity);
    this.outSize = capacity;
  }

  private grow(needed: number): void {
    let newSize = this.outSize + (this.outSize >>> 1);
    if (newSize < needed) newSize = needed + (this.outSize >>> 1);
    if (newSize > MAX_OUT) {
      throw new Error("LZCOMP output exceeds maximum size budget");
    }
    const tmp = new Uint8Array(newSize);
    tmp.set(this.out);
    this.out = tmp;
    this.outSize = newSize;
  }

  push(byte: number): void {
    switch (this.state) {
      case 0:
        this.escape = byte;
        this.state = 1;
        break;
      case 1:
        if (byte === this.escape) {
          this.state = 2;
        } else {
          if (this.outIdx >= this.outSize) this.grow(this.outIdx + 1);
          this.out[this.outIdx++] = byte;
        }
        break;
      case 2:
        this.count = byte;
        if (this.count === 0) {
          if (this.outIdx >= this.outSize) this.grow(this.outIdx + 1);
          this.out[this.outIdx++] = this.escape;
          this.state = 1;
        } else {
          this.state = 3;
        }
        break;
      default: {
        if (this.outIdx + this.count > this.outSize) this.grow(this.outIdx + this.count);
        this.out.fill(byte, this.outIdx, this.outIdx + this.count);
        this.outIdx += this.count;
        this.state = 1;
        break;
      }
    }
  }

  result(): Uint8Array {
    return this.out.subarray(0, this.outIdx);
  }
}

export function lzcompDecompress(data: Uint8Array, size: number, version: number): Uint8Array {
  const bio = new BitReader(data, 0, size);
  const usingRunLength = version === 1 ? false : bio.inputBit() !== 0;

  const distEcoder = new AHuff(bio, 1 << DIST_WIDTH);
  const lenEcoder = new AHuff(bio, 1 << LEN_WIDTH);
  const outLen = bio.readValue(24);
  if (outLen > MAX_OUT_LEN) {
    throw new Error(`LZCOMP outLen ${outLen} exceeds maximum (${MAX_OUT_LEN})`);
  }
  const { DUP2, DUP4, DUP6, NUM_SYMS } = setDistRange(outLen);
  const symEcoder = new AHuff(bio, NUM_SYMS);

  const win = new Uint8Array(PRELOAD_SIZE + outLen);
  initializeModel(win);
  const base = PRELOAD_SIZE;

  // Every emitted byte also lands in the window at base+pos, so in the
  // common non-RLE case the output IS the window tail — no second buffer.
  const rle = usingRunLength ? new RleSink(outLen) : null;

  for (let pos = 0; pos < outLen; ) {
    const symbol = symEcoder.readSymbol();
    let value: number;
    if (symbol < 256) {
      value = symbol;
    } else if (symbol === DUP2) {
      value = win[base + pos - 2];
    } else if (symbol === DUP4) {
      value = win[base + pos - 4];
    } else if (symbol === DUP6) {
      value = win[base + pos - 6];
    } else {
      let length = decodeLength(lenEcoder, symbol);
      const distance = decodeDistance(distEcoder, lastNumDistRanges);
      if (distance >= MAX_2BYTE_DIST) {
        length++;
      }
      const start = base + pos - distance - length + 1;
      // Clamp: a malformed stream could declare a match extending past
      // outLen; never emit more than the declared output size.
      if (length > outLen - pos) length = outLen - pos;
      if (rle) {
        for (let j = 0; j < length; j++) {
          const v = win[start + j];
          win[base + pos] = v;
          pos++;
          rle.push(v);
        }
      } else {
        // Byte-by-byte: source and destination may overlap (LZ semantics).
        for (let j = 0; j < length; j++) {
          win[base + pos] = win[start + j];
          pos++;
        }
      }
      continue;
    }
    win[base + pos] = value;
    pos++;
    if (rle) rle.push(value);
  }

  return rle ? rle.result() : win.subarray(base, base + outLen);
}
