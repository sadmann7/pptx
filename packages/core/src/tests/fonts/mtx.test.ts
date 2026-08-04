/**
 * Round-trip coverage for the MTX decode path.
 *
 * Every embedded font shipped in this repo is an *uncompressed* EOT, so the
 * LZCOMP and CTF layers would otherwise never run. These tests generate
 * compressed fixtures from the format spec (see `mtx-encoder.ts`) and assert
 * the decoder reconstructs the outlines and hints that went in.
 */

import { describe, expect, it } from "vitest";

import { decodeEmbeddedFont } from "../../fonts/decode";
import { decodeMtx, MtxError, parseEotMetadata } from "../../fonts/mtx";
import { Reader } from "../../fonts/mtx/binary";
import { DEFAULT_LIMITS } from "../../fonts/mtx/limits";
import { AdaptiveHuffman, BitReader, decompressLzcomp } from "../../fonts/mtx/lzcomp";
import { decodeTripletArrays } from "../../fonts/mtx/triplet";
import {
  buildCompressedEot,
  buildCtfStreams,
  buildEot,
  buildMtx,
  encodeLzcompGreedy,
  encodeLzcompLiterals,
  encodeLzcompOps,
  encodeTriplet,
  packRunLength,
  TTEMBED_TTCOMPRESSED,
  TTEMBED_XORENCRYPTDATA,
  type CtfFontSpec,
  type GlyphPoint,
} from "./mtx-encoder";
import {
  decodePushBurst,
  TtfReader,
  type ReadCompositeGlyph,
  type ReadSimpleGlyph,
} from "./ttf-reader";

const LIMITS = { ...DEFAULT_LIMITS };

function pt(x: number, y: number, onCurve = true): GlyphPoint {
  return { x, y, onCurve };
}

describe("LZCOMP", () => {
  it("round-trips a literal-only stream", () => {
    const expected = new TextEncoder().encode("MicroType Express literal stream");
    expect(decompressLzcomp(encodeLzcompLiterals(expected), 3, LIMITS)).toEqual(expected);
  });

  it("round-trips back-references produced by a greedy matcher", () => {
    const text =
      "abcabcabcabc the quick brown fox the quick brown fox jumps over the quick brown fox";
    const expected = new TextEncoder().encode(text.repeat(4));
    const encoded = encodeLzcompGreedy(expected);
    expect(encoded.length).toBeLessThan(expected.length);
    expect(decompressLzcomp(encoded, 3, LIMITS)).toEqual(expected);
  });

  it("round-trips a stream long enough to need multiple distance ranges", () => {
    const expected = new Uint8Array(6000);
    for (let i = 0; i < expected.length; i++) expected[i] = (i * 7 + (i >>> 5)) & 0xff;
    expect(decompressLzcomp(encodeLzcompGreedy(expected), 3, LIMITS)).toEqual(expected);
  });

  it("decodes the DUP2, DUP4, and DUP6 shorthands", () => {
    const seed = [1, 2, 3, 4, 5, 6];
    const ops = [
      ...seed.map((literal) => ({ literal })),
      { dup: 2 as const }, // repeats 5
      { dup: 4 as const }, // repeats 4
      { dup: 6 as const }, // repeats 3
    ];
    const decoded = decompressLzcomp(encodeLzcompOps(ops, 9), 3, LIMITS);
    expect(Array.from(decoded)).toEqual([1, 2, 3, 4, 5, 6, 5, 4, 3]);
  });

  it("decodes an explicit near copy", () => {
    const ops = [
      { literal: 0x41 },
      { literal: 0x42 },
      { literal: 0x43 },
      { distance: 1, length: 3 },
    ];
    expect(Array.from(decompressLzcomp(encodeLzcompOps(ops, 6), 3, LIMITS))).toEqual([
      0x41, 0x42, 0x43, 0x41, 0x42, 0x43,
    ]);
  });

  it("expands the optional run-length layer", () => {
    const expected = Uint8Array.from([1, 1, 1, 1, 1, 1, 2, 0xee, 0xee, 3]);
    const packed = packRunLength(expected, 0xee);
    expect(decompressLzcomp(encodeLzcompLiterals(packed, { runLength: true }), 3, LIMITS)).toEqual(
      expected,
    );
  });

  it("reads a version-1 stream, which omits the run-length flag bit", () => {
    const expected = new TextEncoder().encode("version one has no leading flag bit");
    const stream = encodeLzcompGreedy(expected, { version: 1 });
    expect(decompressLzcomp(stream, 1, LIMITS)).toEqual(expected);
    // Reading the same bytes as version 3 consumes a length bit as the flag
    // and desynchronizes the whole stream.
    expect(() => decompressLzcomp(stream, 3, LIMITS)).toThrow(MtxError);
  });

  it("rejects a stream that declares more output than the limit allows", () => {
    const stream = encodeLzcompLiterals(Uint8Array.from([1, 2, 3]));
    expect(() => decompressLzcomp(stream, 3, { ...LIMITS, maxStreamBytes: 2 })).toThrow(
      /declares 3 bytes/u,
    );
  });

  it("keeps the adaptive tree in lockstep between encoder and decoder", () => {
    const tree = new AdaptiveHuffman(8);
    const mirror = new AdaptiveHuffman(8);
    const symbols = [3, 3, 7, 0, 1, 3, 7, 7, 2, 5, 5, 5, 0];
    const bits: number[] = [];
    for (const symbol of symbols) bits.push(...tree.encode(symbol));

    const packed = new Uint8Array(Math.ceil(bits.length / 8));
    bits.forEach((bit, i) => {
      packed[i >>> 3]! |= bit << (7 - (i & 7));
    });
    const reader = new BitReader(packed);
    expect(symbols.map(() => mirror.read(reader))).toEqual(symbols);
  });
});

describe("triplet encoding", () => {
  it("round-trips deltas across every encoding class", () => {
    const deltas: [number, number][] = [
      [0, 0],
      [0, 5],
      [0, -5],
      [7, 0],
      [-7, 0],
      [0, 1024],
      [0, -1279],
      [1024, 0],
      [3, 4],
      [-3, -4],
      [49, 64],
      [-49, 64],
      [1, 1],
      [-1, -1],
      [257, 257],
      [-513, 768],
      [600, -600],
      [1000, 2000],
      [-4095, 4095],
      [4096, -4096],
      [20000, -20000],
      [-32768, 32767],
    ];

    const flags: number[] = [];
    const coordinates: number[] = [];
    for (const [dx, dy] of deltas) {
      const encoded = encodeTriplet(dx, dy, true);
      flags.push(encoded[0]!);
      coordinates.push(...encoded.slice(1));
    }

    const decoded = decodeTripletArrays(
      new Reader(Uint8Array.from(coordinates)),
      Uint8Array.from(flags),
    );
    let x = 0;
    let y = 0;
    deltas.forEach(([dx, dy], i) => {
      x += dx;
      y += dy;
      expect([decoded.x[i], decoded.y[i]]).toEqual([x, y]);
    });
  });

  it("prefers compact encodings for small deltas", () => {
    expect(encodeTriplet(0, 5, true)).toHaveLength(2);
    expect(encodeTriplet(3, 4, true)).toHaveLength(2);
    expect(encodeTriplet(300, 300, true)).toHaveLength(3);
    expect(encodeTriplet(20000, -20000, true)).toHaveLength(5);
  });

  it("carries the off-curve bit", () => {
    expect(encodeTriplet(3, 4, false)[0]! & 0x80).toBe(0x80);
    expect(encodeTriplet(3, 4, true)[0]! & 0x80).toBe(0);
  });
});

// A font exercising each CTF glyph shape: contourless, computed bounding box,
// explicit bounding box, composite, and hop-coded hint pushes.
const SQUARE: GlyphPoint[] = [pt(50, 50), pt(650, 50), pt(650, 700), pt(50, 700)];
const INNER: GlyphPoint[] = [pt(150, 150), pt(300, 200, false), pt(550, 150), pt(550, 600, false)];
const WIDE: GlyphPoint[] = [
  pt(-400, -250),
  pt(1800, -250),
  pt(1800, 1900),
  pt(-400, 1900),
  pt(700, 2600, false),
];

const FONT: CtfFontSpec = {
  cvt: [0, 100, 337, 575, 574, 274, 65000, 1200, 1200],
  glyphs: [
    { kind: "simple", contours: [] },
    { kind: "simple", contours: [SQUARE, INNER] },
    {
      kind: "simple",
      contours: [WIDE],
      bbox: [-500, -300, 1900, 2700],
      pushes: [1, 2, 300, -400, 0, 255],
      code: [0x2f, 0x18, 0x21],
    },
    {
      kind: "composite",
      bbox: [0, 0, 1400, 800],
      components: [
        { flags: 0x0002, glyphIndex: 1, args: [0, 0] },
        { flags: 0x0102, glyphIndex: 1, args: [100, 20] },
      ],
      pushes: [10, 20],
      code: [0x2c],
    },
    {
      kind: "simple",
      contours: [[pt(0, 0), pt(100, 0), pt(100, 100)]],
      pushes: [7, 500, 7, -600, 7, 9, 9, 800, 9],
      useHopCodes: true,
    },
  ],
};

function decodeFont(spec: CtfFontSpec): TtfReader {
  const decoded = decodeEmbeddedFont(buildCompressedEot(spec));
  expect(decoded).toBeDefined();
  return new TtfReader(decoded!);
}

describe("CTF reconstruction", () => {
  const font = decodeFont(FONT);

  it("produces a valid sfnt with a binary-sorted table directory", () => {
    const tags = font.tableTags;
    expect(tags).toEqual([...tags].sort());
    expect(new Set(tags)).toEqual(new Set(["head", "maxp", "cvt ", "glyf", "loca"]));
    expect(font.numGlyphs).toBe(FONT.glyphs.length);
  });

  it("zeroes head.checksumAdjustment before recomputing it", () => {
    // The encoder seeds it with 0xDEADBEEF; a correct decoder must discard
    // that before the whole-font checksum can balance.
    const head = font.tables.get("head")!;
    const view = new DataView(head.buffer, head.byteOffset, head.byteLength);
    expect(view.getUint32(8, false)).not.toBe(0xdeadbeef);
    expect(view.getUint16(18, false)).toBe(1000);
  });

  it("gives a contourless glyph an empty loca range", () => {
    const offsets = font.locaOffsets;
    expect(offsets[0]).toBe(0);
    expect(offsets[1]).toBe(0);
    expect(font.glyph(0)).toEqual({ kind: "empty" });
  });

  it("rebuilds contours and computes the bounding box when it is omitted", () => {
    const glyph = font.glyph(1) as ReadSimpleGlyph;
    expect(glyph.kind).toBe("simple");
    expect(glyph.contours).toEqual([SQUARE, INNER]);
    expect(glyph.bbox).toEqual([50, 50, 650, 700]);
  });

  it("honours an explicit bounding box and reassembles hints", () => {
    const glyph = font.glyph(2) as ReadSimpleGlyph;
    expect(glyph.contours).toEqual([WIDE]);
    expect(glyph.bbox).toEqual([-500, -300, 1900, 2700]);

    const { values, rest } = decodePushBurst(glyph.instructions);
    expect(values).toEqual([1, 2, 300, -400, 0, 255]);
    expect(rest).toEqual([0x2f, 0x18, 0x21]);
  });

  it("rebuilds composite components and their instructions", () => {
    const glyph = font.glyph(3) as ReadCompositeGlyph;
    expect(glyph.kind).toBe("composite");
    expect(glyph.bbox).toEqual([0, 0, 1400, 800]);
    expect(glyph.components.map((c) => c.glyphIndex)).toEqual([1, 1]);
    expect(glyph.components.map((c) => c.args)).toEqual([
      [0, 0],
      [100, 20],
    ]);
    expect(glyph.components[0]!.flags & 0x0020).toBe(0x0020);
    expect(glyph.components[1]!.flags & 0x0020).toBe(0);

    const { values, rest } = decodePushBurst(glyph.instructions);
    expect(values).toEqual([10, 20]);
    expect(rest).toEqual([0x2c]);
  });

  it("expands hop codes in the push stream", () => {
    const glyph = font.glyph(4) as ReadSimpleGlyph;
    const { values } = decodePushBurst(glyph.instructions);
    expect(values).toEqual([7, 500, 7, -600, 7, 9, 9, 800, 9]);
  });

  it("decodes the control value table as an entry count, not a byte length", () => {
    const cvt = font.tables.get("cvt ")!;
    const view = new DataView(cvt.buffer, cvt.byteOffset, cvt.byteLength);
    const values = Array.from({ length: cvt.length / 2 }, (_, i) => view.getUint16(i * 2, false));
    expect(values).toEqual(FONT.cvt);
  });

  it("switches loca to the long format when glyf outgrows short offsets", () => {
    // Short offsets store glyf/2 in a USHORT, so anything past 131070 bytes
    // has to be rewritten as long even though head asked for short.
    const filler = Array.from({ length: 250 }, (_, i) => ({
      kind: "simple" as const,
      contours: [
        Array.from({ length: 250 }, (_, j) =>
          pt(1000 + (((i + j) * 137) % 5000), 500 + (((i + j) * 211) % 4000)),
        ),
      ],
    }));
    const decoded = decodeEmbeddedFont(
      buildCompressedEot({ glyphs: filler, indexToLocFormat: 0 }, { literalOnly: true }),
    );
    expect(decoded).toBeDefined();
    const big = new TtfReader(decoded!);
    const glyf = big.tables.get("glyf")!;

    expect(glyf.length).toBeGreaterThan(0xffff * 2);
    expect(big.indexToLocFormat).toBe(1);
    expect(big.locaOffsets.at(-1)).toBe(glyf.length);
  });
});

describe("MTX container", () => {
  it("accepts a version-1 container", () => {
    const eot = buildCompressedEot(FONT, { version: 1 });
    expect(new TtfReader(decodeEmbeddedFont(eot)!).numGlyphs).toBe(FONT.glyphs.length);
  });

  it("rejects an unknown container version", () => {
    const streams = buildCtfStreams(FONT).map((stream) => encodeLzcompGreedy(stream));
    const mtx = buildMtx(streams as [Uint8Array, Uint8Array, Uint8Array], 2);
    expect(() => decodeMtx(mtx)).toThrow(/Unsupported MTX version 2/u);
  });

  it("undoes XOR obfuscation of the payload", () => {
    const eot = buildCompressedEot(FONT, { encrypted: true });
    expect(parseEotMetadata(eot).encrypted).toBe(true);
    expect(new TtfReader(decodeEmbeddedFont(eot)!).numGlyphs).toBe(FONT.glyphs.length);
  });
});

describe("EOT container", () => {
  it("reads the compression and encryption flags from the header", () => {
    const metadata = parseEotMetadata(buildEot(Uint8Array.of(1, 2, 3, 4), TTEMBED_TTCOMPRESSED));
    expect(metadata).toMatchObject({
      version: 1,
      fontDataSize: 4,
      compressed: true,
      encrypted: false,
    });
    expect(metadata.fontDataOffset).toBe(metadata.totalSize - 4);
  });

  it("reports both flags when set together", () => {
    const flags = TTEMBED_TTCOMPRESSED | TTEMBED_XORENCRYPTDATA;
    expect(parseEotMetadata(buildEot(Uint8Array.of(9), flags))).toMatchObject({
      compressed: true,
      encrypted: true,
    });
  });

  it("rejects a header without the EOT magic number", () => {
    const eot = buildEot(Uint8Array.of(1, 2, 3, 4), 0);
    new DataView(eot.buffer).setUint16(34, 0x0000, true);
    expect(() => parseEotMetadata(eot)).toThrow(/magic number/u);
  });

  it("passes an uncompressed payload straight through", () => {
    const streams = buildCtfStreams(FONT);
    const font = decodeMtx(streams[0], { compressed: false });
    expect(font).toEqual(streams[0]);
  });
});
