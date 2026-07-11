/**
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * @remarks
 * Derived from mtx-decompressor v1.4.2
 * (https://github.com/ChristopherVR/mtx-decompressor, © ChristopherVR,
 * MPL-2.0). Optimized for @diceui/pptx-core: geometric buffer growth and
 * bulk copies; behavior is byte-identical to the original.
 */

/** Growable big-endian byte stream with bit-level reads for triplet decoding. */
export class MtxStream {
  buf: Uint8Array;
  /** How much data has been written or is valid. */
  size: number;
  /** Allocated capacity. */
  reserved: number;
  /** Current byte position. */
  pos: number;
  /** Current bit position within the byte at `pos`. */
  bitPos: number;

  constructor(buf: Uint8Array | null, size: number) {
    if (buf) {
      this.buf = buf;
      this.size = size;
      this.reserved = buf.length;
    } else {
      this.buf = new Uint8Array(0);
      this.size = 0;
      this.reserved = 0;
    }
    this.pos = 0;
    this.bitPos = 0;
  }

  reserve(n: number): void {
    if (this.reserved >= n) return;
    const newBuf = new Uint8Array(n);
    newBuf.set(this.buf.subarray(0, this.size));
    this.buf = newBuf;
    this.reserved = n;
  }

  /** Reserve with geometric growth so repeated small reserves stay O(n). */
  reserveGrow(n: number): void {
    if (this.reserved >= n) return;
    this.reserve(Math.max(n, this.reserved * 2 || 256));
  }

  ensureWrite(n: number): void {
    const needed = this.pos + n;
    if (needed > this.reserved) {
      this.reserve(Math.max(needed, this.reserved * 2 || 256));
    }
    if (needed > this.size) {
      this.size = needed;
    }
  }

  ensureRead(n: number): void {
    if (this.pos + n > this.size) {
      throw new Error(
        `Stream: not enough data (need ${n} bytes at pos ${this.pos}, size ${this.size})`,
      );
    }
  }

  seekAbsolute(pos: number): void {
    if (pos > this.size) {
      throw new Error(`Stream: seek past end (${pos} > ${this.size})`);
    }
    this.pos = pos;
    this.bitPos = 0;
  }

  seekRelative(offset: number): void {
    const newPos = this.pos + offset;
    if (newPos < 0) throw new Error("Stream: negative seek");
    if (newPos > this.size) throw new Error("Stream: seek past end");
    this.pos = newPos;
    this.bitPos = 0;
  }

  // --- Read (big-endian) ---

  readU8(): number {
    this.ensureRead(1);
    return this.buf[this.pos++];
  }

  peekU8(): number {
    this.ensureRead(1);
    return this.buf[this.pos];
  }

  readU16(): number {
    this.ensureRead(2);
    const v = (this.buf[this.pos] << 8) | this.buf[this.pos + 1];
    this.pos += 2;
    return v;
  }

  readU32(): number {
    this.ensureRead(4);
    const v =
      ((this.buf[this.pos] << 24) |
        (this.buf[this.pos + 1] << 16) |
        (this.buf[this.pos + 2] << 8) |
        this.buf[this.pos + 3]) >>>
      0;
    this.pos += 4;
    return v;
  }

  readS16(): number {
    const v = this.readU16();
    return v >= 32768 ? v - 65536 : v;
  }

  // --- Write (big-endian) ---

  writeU8(v: number): void {
    this.ensureWrite(1);
    this.buf[this.pos++] = v & 255;
  }

  writeU16(v: number): void {
    this.ensureWrite(2);
    this.buf[this.pos++] = (v >> 8) & 255;
    this.buf[this.pos++] = v & 255;
  }

  writeU32(v: number): void {
    this.ensureWrite(4);
    this.buf[this.pos++] = (v >>> 24) & 255;
    this.buf[this.pos++] = (v >> 16) & 255;
    this.buf[this.pos++] = (v >> 8) & 255;
    this.buf[this.pos++] = v & 255;
  }

  writeS16(v: number): void {
    this.writeU16(v < 0 ? v + 65536 : v);
  }

  // --- Bit-level reading (triplet coordinate decoding) ---

  readNBits(n: number): number {
    if (n === 0) return 0;
    let value = 0;
    let bitsRemaining = n;
    while (bitsRemaining > 0) {
      if (this.pos >= this.size && this.bitPos === 0) {
        throw new Error("Stream: not enough data for bit read");
      }
      const bitsAvailableInByte = 8 - this.bitPos;
      const bitsToRead = bitsRemaining < bitsAvailableInByte ? bitsRemaining : bitsAvailableInByte;
      const shift = bitsAvailableInByte - bitsToRead;
      const mask = ((1 << bitsToRead) - 1) << shift;
      value = (value << bitsToRead) | ((this.buf[this.pos] & mask) >> shift);
      this.bitPos += bitsToRead;
      if (this.bitPos >= 8) {
        this.bitPos = 0;
        this.pos++;
      }
      bitsRemaining -= bitsToRead;
    }
    return value;
  }

  // --- Copy ---

  /** Copy `length` bytes from this stream to `dest`. */
  copyTo(dest: MtxStream, length: number): void {
    if (this.pos + length > this.size) {
      throw new Error("Stream: not enough data for copy");
    }
    dest.ensureWrite(length);
    dest.buf.set(this.buf.subarray(this.pos, this.pos + length), dest.pos);
    this.pos += length;
    dest.pos += length;
  }

  /** Get a copy of the written data. */
  toUint8Array(): Uint8Array {
    return this.buf.slice(0, this.size);
  }
}
