/**
 * Test-only TrueType reader.
 *
 * Just enough of the sfnt/glyf format to assert that a decoded font carries
 * the outlines and hints the fixture encoder put in, without asserting on the
 * exact byte encoding the decoder chose.
 */

export interface ReadPoint {
  x: number;
  y: number;
  onCurve: boolean;
}

export interface ReadSimpleGlyph {
  kind: "simple";
  bbox: [number, number, number, number];
  contours: ReadPoint[][];
  instructions: number[];
}

export interface ReadCompositeGlyph {
  kind: "composite";
  bbox: [number, number, number, number];
  components: { flags: number; glyphIndex: number; args: number[] }[];
  instructions: number[];
}

export type ReadGlyph = ReadSimpleGlyph | ReadCompositeGlyph | { kind: "empty" };

export class TtfReader {
  readonly tables = new Map<string, Uint8Array>();
  private readonly view: DataView;

  constructor(private readonly font: Uint8Array) {
    this.view = new DataView(font.buffer, font.byteOffset, font.byteLength);
    const count = this.view.getUint16(4, false);
    for (let i = 0; i < count; i++) {
      const record = 12 + i * 16;
      const tag = String.fromCharCode(
        font[record]!,
        font[record + 1]!,
        font[record + 2]!,
        font[record + 3]!,
      );
      const offset = this.view.getUint32(record + 8, false);
      const length = this.view.getUint32(record + 12, false);
      this.tables.set(tag, font.subarray(offset, offset + length));
    }
  }

  get tableTags(): string[] {
    const count = this.view.getUint16(4, false);
    return Array.from({ length: count }, (_, i) => {
      const record = 12 + i * 16;
      return String.fromCharCode(
        this.font[record]!,
        this.font[record + 1]!,
        this.font[record + 2]!,
        this.font[record + 3]!,
      );
    });
  }

  private table(tag: string): Uint8Array {
    const data = this.tables.get(tag);
    if (!data) throw new Error(`Font has no ${tag} table`);
    return data;
  }

  get numGlyphs(): number {
    const maxp = this.table("maxp");
    return new DataView(maxp.buffer, maxp.byteOffset, maxp.byteLength).getUint16(4, false);
  }

  get indexToLocFormat(): number {
    const head = this.table("head");
    return new DataView(head.buffer, head.byteOffset, head.byteLength).getInt16(50, false);
  }

  get locaOffsets(): number[] {
    const loca = this.table("loca");
    const view = new DataView(loca.buffer, loca.byteOffset, loca.byteLength);
    const short = this.indexToLocFormat === 0;
    const count = this.numGlyphs + 1;
    return Array.from({ length: count }, (_, i) =>
      short ? view.getUint16(i * 2, false) * 2 : view.getUint32(i * 4, false),
    );
  }

  glyph(index: number): ReadGlyph {
    const offsets = this.locaOffsets;
    const start = offsets[index]!;
    const end = offsets[index + 1]!;
    if (end <= start) return { kind: "empty" };

    const glyf = this.table("glyf").subarray(start, end);
    const view = new DataView(glyf.buffer, glyf.byteOffset, glyf.byteLength);
    const numContours = view.getInt16(0, false);
    const bbox: [number, number, number, number] = [
      view.getInt16(2, false),
      view.getInt16(4, false),
      view.getInt16(6, false),
      view.getInt16(8, false),
    ];

    return numContours < 0
      ? readComposite(glyf, view, bbox)
      : readSimple(glyf, view, bbox, numContours);
  }
}

function readSimple(
  glyf: Uint8Array,
  view: DataView,
  bbox: [number, number, number, number],
  numContours: number,
): ReadSimpleGlyph {
  const endPoints = Array.from({ length: numContours }, (_, i) =>
    view.getUint16(10 + i * 2, false),
  );
  const pointCount = numContours === 0 ? 0 : endPoints[numContours - 1]! + 1;
  let pos = 10 + numContours * 2;
  const instructionLength = view.getUint16(pos, false);
  pos += 2;
  const instructions = Array.from(glyf.subarray(pos, pos + instructionLength));
  pos += instructionLength;

  const flags: number[] = [];
  while (flags.length < pointCount) {
    const flag = glyf[pos++]!;
    flags.push(flag);
    if (flag & 0x08) {
      const repeat = glyf[pos++]!;
      for (let i = 0; i < repeat; i++) flags.push(flag);
    }
  }
  if (flags.length !== pointCount) throw new Error("Flag run overshot the point count");

  const xs: number[] = [];
  let x = 0;
  for (const flag of flags) {
    if (flag & 0x02) {
      const delta = glyf[pos++]!;
      x += flag & 0x10 ? delta : -delta;
    } else if (!(flag & 0x10)) {
      x += view.getInt16(pos, false);
      pos += 2;
    }
    xs.push(x);
  }

  const ys: number[] = [];
  let y = 0;
  for (const flag of flags) {
    if (flag & 0x04) {
      const delta = glyf[pos++]!;
      y += flag & 0x20 ? delta : -delta;
    } else if (!(flag & 0x20)) {
      y += view.getInt16(pos, false);
      pos += 2;
    }
    ys.push(y);
  }

  const contours: ReadPoint[][] = [];
  let cursor = 0;
  for (const endPoint of endPoints) {
    const contour: ReadPoint[] = [];
    for (; cursor <= endPoint; cursor++) {
      contour.push({ x: xs[cursor]!, y: ys[cursor]!, onCurve: (flags[cursor]! & 0x01) !== 0 });
    }
    contours.push(contour);
  }

  return { kind: "simple", bbox, contours, instructions };
}

function readComposite(
  glyf: Uint8Array,
  view: DataView,
  bbox: [number, number, number, number],
): ReadCompositeGlyph {
  const components: ReadCompositeGlyph["components"] = [];
  let pos = 10;
  let flags = 0x0020;
  while (flags & 0x0020) {
    flags = view.getUint16(pos, false);
    const glyphIndex = view.getUint16(pos + 2, false);
    pos += 4;
    const argBytes = flags & 0x0001 ? 4 : 2;
    const args = Array.from(glyf.subarray(pos, pos + argBytes));
    pos += argBytes;
    if (flags & 0x0080) pos += 8;
    else if (flags & 0x0040) pos += 4;
    else if (flags & 0x0008) pos += 2;
    components.push({ flags, glyphIndex, args });
  }

  let instructions: number[] = [];
  if (flags & 0x0100) {
    const length = view.getUint16(pos, false);
    instructions = Array.from(glyf.subarray(pos + 2, pos + 2 + length));
  }
  return { kind: "composite", bbox, components, instructions };
}

/** Expand a TrueType PUSH burst back into the values it pushes. */
export function decodePushBurst(instructions: readonly number[]): {
  values: number[];
  rest: number[];
} {
  const values: number[] = [];
  let pos = 0;
  for (;;) {
    const op = instructions[pos];
    if (op === undefined) break;
    let count: number;
    let words: boolean;
    if (op === 0x40) {
      count = instructions[pos + 1]!;
      words = false;
      pos += 2;
    } else if (op === 0x41) {
      count = instructions[pos + 1]!;
      words = true;
      pos += 2;
    } else if (op >= 0xb0 && op <= 0xb7) {
      count = op - 0xb0 + 1;
      words = false;
      pos += 1;
    } else if (op >= 0xb8 && op <= 0xbf) {
      count = op - 0xb8 + 1;
      words = true;
      pos += 1;
    } else break;

    for (let i = 0; i < count; i++) {
      if (words) {
        const high = instructions[pos]!;
        const low = instructions[pos + 1]!;
        const value = (high << 8) | low;
        values.push(value & 0x8000 ? value - 0x10000 : value);
        pos += 2;
      } else {
        values.push(instructions[pos]!);
        pos += 1;
      }
    }
  }
  return { values, rest: instructions.slice(pos) };
}
