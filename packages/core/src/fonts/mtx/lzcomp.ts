import { fail } from "./error";
import type { DecoderLimits } from "./limits";

const PRELOAD_SIZE = 2 * 32 * 96 + 4 * 256;

/**
 * MSB-first bit reader with a byte-at-a-time prefetch buffer.
 *
 * `buffer` and `available` are public so the Huffman descent can hold them in
 * locals for a whole symbol instead of paying a call and a bounds check per
 * bit. Prefetching whole bytes is safe: it only fails once a bit is actually
 * requested past the end of the input.
 */
export class BitReader {
  private readonly data: Uint8Array;
  private index = 0;
  /** Low `available` bits are valid, MSB-first (the next bit is the highest). */
  buffer = 0;
  available = 0;

  constructor(data: Uint8Array) {
    this.data = data;
  }

  /** Load bytes until at least one bit is available. */
  refill(): void {
    if (this.index >= this.data.length) fail("BOUNDS", "Unexpected end of LZCOMP bitstream");
    do {
      this.buffer = ((this.buffer << 8) | this.data[this.index++]!) >>> 0;
      this.available += 8;
    } while (this.available <= 16 && this.index < this.data.length);
  }

  bit(): number {
    if (this.available === 0) this.refill();
    return (this.buffer >>> --this.available) & 1;
  }

  bits(count: number): number {
    let value = 0;
    for (let i = 0; i < count; i++) value = value * 2 + this.bit();
    return value;
  }
}

/** Array-backed adaptive Huffman tree from the MTX format specification. */
export class AdaptiveHuffman {
  private readonly range: number;
  private readonly up: Int32Array;
  /** Children interleaved as `child[node * 2 + bit]`, so a descent step indexes
   * instead of branching on a bit that is inherently unpredictable. */
  private readonly child: Int32Array;
  private readonly code: Int32Array;
  private readonly weight: Int32Array;
  private readonly symbolIndex: Int32Array;

  constructor(range: number) {
    if (!Number.isInteger(range) || range < 2 || range >= 512) {
      fail("INVALID_MTX", `Invalid adaptive-Huffman range ${range}`);
    }
    this.range = range;
    const nodeCount = range * 2;
    this.up = new Int32Array(nodeCount);
    this.child = new Int32Array(nodeCount * 2);
    this.code = new Int32Array(nodeCount);
    this.weight = new Int32Array(nodeCount);
    this.symbolIndex = new Int32Array(range);

    this.code.fill(-1);
    this.child.fill(-1);
    for (let i = 2; i < nodeCount; i++) {
      this.up[i] = i >>> 1;
      this.weight[i] = 1;
    }
    for (let i = 1; i < range; i++) {
      this.child[i * 2] = i * 2;
      this.child[i * 2 + 1] = i * 2 + 1;
    }
    for (let symbol = 0; symbol < range; symbol++) {
      const node = range + symbol;
      this.code[node] = symbol;
      this.symbolIndex[symbol] = node;
    }
    for (let i = range - 1; i >= 1; i--) {
      this.weight[i] = this.weight[this.child[i * 2]!]! + this.weight[this.child[i * 2 + 1]!]!;
    }

    if (range > 256) {
      this.update(this.symbolIndex[256]!);
      this.update(this.symbolIndex[257]!);
      for (let i = 0; i < 12; i++) this.update(this.symbolIndex[range - 3]!);
      for (let i = 0; i < 6; i++) this.update(this.symbolIndex[range - 2]!);
    } else {
      for (let pass = 0; pass < 2; pass++) {
        for (let symbol = 0; symbol < range; symbol++) this.update(this.symbolIndex[symbol]!);
      }
    }
  }

  read(bits: BitReader): number {
    // The root is never a leaf, so the descent can read before it tests.
    const { child, code } = this;
    let buffer = bits.buffer;
    let available = bits.available;
    let node = 1;
    let symbol: number;
    do {
      if (available === 0) {
        bits.buffer = buffer;
        bits.available = available;
        bits.refill();
        buffer = bits.buffer;
        available = bits.available;
      }
      node = child[node * 2 + ((buffer >>> --available) & 1)]!;
      symbol = code[node]!;
    } while (symbol < 0);
    bits.buffer = buffer;
    bits.available = available;
    this.update(node);
    return symbol;
  }

  /** Returns the current code bits, then updates the model. Used by tests. */
  encode(symbol: number): number[] {
    if (symbol < 0 || symbol >= this.range) fail("INVALID_MTX", `Symbol ${symbol} outside tree`);
    let node = this.symbolIndex[symbol]!;
    const reverse: number[] = [];
    const original = node;
    while (node !== 1) {
      const parent = this.up[node]!;
      reverse.push(this.child[parent * 2 + 1] === node ? 1 : 0);
      node = parent;
    }
    this.update(original);
    reverse.reverse();
    return reverse;
  }

  /** Exchange two nodes' contents; `up` pointers stay with the positions. */
  private swap(a: number, b: number): void {
    const { child, code, weight } = this;
    let tmp = child[a * 2]!;
    child[a * 2] = child[b * 2]!;
    child[b * 2] = tmp;
    tmp = child[a * 2 + 1]!;
    child[a * 2 + 1] = child[b * 2 + 1]!;
    child[b * 2 + 1] = tmp;
    tmp = code[a]!;
    code[a] = code[b]!;
    code[b] = tmp;
    tmp = weight[a]!;
    weight[a] = weight[b]!;
    weight[b] = tmp;
    this.repair(a);
    this.repair(b);
  }

  private repair(node: number): void {
    const symbol = this.code[node]!;
    if (symbol < 0) {
      this.up[this.child[node * 2]!] = node;
      this.up[this.child[node * 2 + 1]!] = node;
    } else {
      this.symbolIndex[symbol] = node;
    }
  }

  private update(start: number): void {
    let node = start;
    while (node !== 1) {
      const oldWeight = this.weight[node]!;
      let peer = node - 1;
      if (this.weight[peer] === oldWeight) {
        do peer--;
        while (this.weight[peer] === oldWeight);
        peer++;
        if (peer > 1) {
          this.swap(node, peer);
          node = peer;
        }
      }
      this.weight[node] = oldWeight + 1;
      node = this.up[node]!;
    }
    this.weight[1] = this.weight[1]! + 1;
  }
}

function makePreload(): Uint8Array {
  const target = new Uint8Array(PRELOAD_SIZE);
  let p = 0;
  for (let k = 0; k < 32; k++) {
    for (let j = 0; j < 96; j++) {
      target[p++] = k;
      target[p++] = j;
    }
  }
  for (let j = 0; j < 256; j++) {
    target[p++] = j;
    target[p++] = j;
    target[p++] = j;
    target[p++] = j;
  }
  return target;
}

const PRELOAD = makePreload();

function expandRunLength(encoded: Uint8Array, limit: number): Uint8Array {
  if (encoded.length === 0) fail("INVALID_MTX", "Run-length stream has no escape byte");
  const escape = encoded[0]!;
  let length = 0;
  for (let i = 1; i < encoded.length;) {
    const value = encoded[i++]!;
    if (value !== escape) {
      length++;
    } else {
      if (i >= encoded.length) fail("INVALID_MTX", "Truncated run-length escape");
      const count = encoded[i++]!;
      if (count === 0) length++;
      else {
        if (i >= encoded.length) fail("INVALID_MTX", "Truncated run-length value");
        i++;
        length += count;
      }
    }
    if (length > limit) fail("LIMIT_EXCEEDED", "Run-length output exceeds configured limit");
  }

  const out = new Uint8Array(length);
  let write = 0;
  for (let i = 1; i < encoded.length;) {
    const value = encoded[i++]!;
    if (value !== escape) out[write++] = value;
    else {
      const count = encoded[i++]!;
      if (count === 0) out[write++] = escape;
      else out.fill(encoded[i++]!, write, (write += count));
    }
  }
  return out;
}

export function decompressLzcomp(
  input: Uint8Array,
  version: number,
  limits: DecoderLimits,
): Uint8Array {
  const bits = new BitReader(input);
  const runLengthEncoded = version === 1 ? false : bits.bit() !== 0;
  const encodedLength = bits.bits(24);
  if (encodedLength > limits.maxStreamBytes) {
    fail("LIMIT_EXCEEDED", `LZCOMP stream declares ${encodedLength} bytes`);
  }
  if (encodedLength === 0) return new Uint8Array();

  // Each range contributes 3 distance bits, so n ranges reach 2^(3n) bytes.
  // The declared length is a 24-bit field, so n = 8 must stay reachable; the
  // resulting alphabet (256 + 8n + 3 = 323) is still within the tree's limit.
  let distanceRanges = 1;
  while (2 ** (3 * distanceRanges) < encodedLength) distanceRanges++;
  if (distanceRanges > 8) fail("INVALID_MTX", "LZCOMP distance range is not representable");
  const dup2 = 256 + 8 * distanceRanges;
  const dup4 = dup2 + 1;
  const dup6 = dup4 + 1;

  const distanceTree = new AdaptiveHuffman(8);
  const lengthTree = new AdaptiveHuffman(8);
  const symbolTree = new AdaptiveHuffman(dup6 + 1);
  const history = new Uint8Array(PRELOAD_SIZE + encodedLength);
  history.set(PRELOAD);
  let produced = 0;

  while (produced < encodedLength) {
    const symbol = symbolTree.read(bits);
    const write = PRELOAD_SIZE + produced;
    if (symbol < 256) {
      history[write] = symbol;
      produced++;
      continue;
    }
    if (symbol === dup2 || symbol === dup4 || symbol === dup6) {
      const distance = symbol === dup2 ? 2 : symbol === dup4 ? 4 : 6;
      history[write] = history[write - distance]!;
      produced++;
      continue;
    }

    let range = symbol - 256;
    const usedDistanceRanges = Math.floor(range / 8) + 1;
    if (usedDistanceRanges > distanceRanges)
      fail("INVALID_MTX", "Invalid LZCOMP distance symbol count");
    range &= 7;
    let length = range & 3;
    while ((range & 4) !== 0) {
      range = lengthTree.read(bits);
      length = length * 4 + (range & 3);
      if (length > encodedLength) fail("INVALID_MTX", "LZCOMP copy length overflow");
    }
    length += 2;

    let distance = 0;
    for (let i = 0; i < usedDistanceRanges; i++) {
      distance = distance * 8 + distanceTree.read(bits);
    }
    distance++;
    if (distance >= 512) length++;
    if (length > encodedLength - produced)
      fail("INVALID_MTX", "LZCOMP copy exceeds declared output");

    const source = PRELOAD_SIZE + produced - distance - length + 1;
    if (source < 0 || source >= write) {
      fail(
        "INVALID_MTX",
        `Invalid LZCOMP copy distance ${distance} for length ${length} at output ${produced}`,
      );
    }
    // The MTX distance is measured from the copied phrase's tail, so the
    // complete source range always ends at or before the destination.
    history.copyWithin(write, source, source + length);
    produced += length;
  }

  const decoded = history.slice(PRELOAD_SIZE);
  return runLengthEncoded ? expandRunLength(decoded, limits.maxExpandedStreamBytes) : decoded;
}
