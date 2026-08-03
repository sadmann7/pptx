import { fail } from "./errors";

export class Reader {
  readonly data: Uint8Array;
  pos: number;
  readonly end: number;

  constructor(data: Uint8Array, start = 0, end = data.length) {
    if (start < 0 || end < start || end > data.length) fail("BOUNDS", "Invalid reader range");
    this.data = data;
    this.pos = start;
    this.end = end;
  }

  get remaining(): number {
    return this.end - this.pos;
  }

  ensure(count: number): void {
    if (!Number.isSafeInteger(count) || count < 0 || this.pos + count > this.end) {
      fail("BOUNDS", `Need ${count} byte(s), only ${this.remaining} remain`, this.pos);
    }
  }

  seek(pos: number): void {
    if (!Number.isSafeInteger(pos) || pos < 0 || pos > this.end)
      fail("BOUNDS", "Invalid seek", pos);
    this.pos = pos;
  }

  skip(count: number): void {
    this.ensure(count);
    this.pos += count;
  }

  u8(): number {
    this.ensure(1);
    return this.data[this.pos++]!;
  }
  i8(): number {
    const n = this.u8();
    return n & 0x80 ? n - 0x100 : n;
  }

  u16be(): number {
    this.ensure(2);
    const p = this.pos;
    this.pos += 2;
    return (this.data[p]! << 8) | this.data[p + 1]!;
  }

  i16be(): number {
    const n = this.u16be();
    return n & 0x8000 ? n - 0x10000 : n;
  }

  u24be(): number {
    this.ensure(3);
    const p = this.pos;
    this.pos += 3;
    return this.data[p]! * 0x10000 + (this.data[p + 1]! << 8) + this.data[p + 2]!;
  }

  u32be(): number {
    this.ensure(4);
    const p = this.pos;
    this.pos += 4;
    return (
      (this.data[p]! * 0x1000000 +
        (this.data[p + 1]! << 16) +
        (this.data[p + 2]! << 8) +
        this.data[p + 3]!) >>>
      0
    );
  }

  u16le(): number {
    this.ensure(2);
    const p = this.pos;
    this.pos += 2;
    return this.data[p]! | (this.data[p + 1]! << 8);
  }

  u32le(): number {
    this.ensure(4);
    const p = this.pos;
    this.pos += 4;
    return (
      (this.data[p]! +
        (this.data[p + 1]! << 8) +
        (this.data[p + 2]! << 16) +
        this.data[p + 3]! * 0x1000000) >>>
      0
    );
  }

  bytes(count: number): Uint8Array {
    this.ensure(count);
    const out = this.data.subarray(this.pos, this.pos + count);
    this.pos += count;
    return out;
  }
}

export class Writer {
  private data: Uint8Array;
  private view: DataView;
  length = 0;

  constructor(capacity = 256) {
    this.data = new Uint8Array(Math.max(1, capacity));
    this.view = new DataView(this.data.buffer);
  }

  private grow(extra: number): void {
    const required = this.length + extra;
    if (required <= this.data.length) return;
    let capacity = this.data.length;
    while (capacity < required) capacity = Math.max(required, capacity * 2);
    const next = new Uint8Array(capacity);
    next.set(this.data.subarray(0, this.length));
    this.data = next;
    this.view = new DataView(next.buffer);
  }

  u8(value: number): void {
    this.grow(1);
    this.data[this.length++] = value;
  }
  i8(value: number): void {
    this.u8(value);
  }
  u16be(value: number): void {
    this.grow(2);
    this.view.setUint16(this.length, value, false);
    this.length += 2;
  }
  i16be(value: number): void {
    this.grow(2);
    this.view.setInt16(this.length, value, false);
    this.length += 2;
  }
  u24be(value: number): void {
    this.u8(value >>> 16);
    this.u8(value >>> 8);
    this.u8(value);
  }
  u32be(value: number): void {
    this.grow(4);
    this.view.setUint32(this.length, value >>> 0, false);
    this.length += 4;
  }

  bytes(value: Uint8Array): void {
    this.grow(value.length);
    this.data.set(value, this.length);
    this.length += value.length;
  }

  zeros(count: number): void {
    this.grow(count);
    this.data.fill(0, this.length, this.length + count);
    this.length += count;
  }
  align(multiple: number): void {
    while (this.length % multiple) this.u8(0);
  }
  patchU16be(offset: number, value: number): void {
    if (offset + 2 > this.length) fail("BOUNDS", "Patch outside output", offset);
    this.view.setUint16(offset, value, false);
  }
  patchU32be(offset: number, value: number): void {
    if (offset + 4 > this.length) fail("BOUNDS", "Patch outside output", offset);
    this.view.setUint32(offset, value >>> 0, false);
  }
  finish(): Uint8Array {
    return this.data.slice(0, this.length);
  }
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
