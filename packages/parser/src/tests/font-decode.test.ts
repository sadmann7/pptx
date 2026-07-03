import { describe, expect, it } from "vitest";

import { decodeEmbeddedFont, toStandaloneArrayBuffer } from "../utils/font-decode";
import { deobfuscateFont } from "../utils/font-deobfuscate";

const TEST_GUID = "{01234567-89AB-CDEF-0123-456789ABCDEF}";

/**
 * Derive the ODTTF XOR key per ECMA-376 Part 2: strip braces/hyphens to get
 * 16 big-endian bytes, then reverse the byte order.
 */
function odttfKey(guid: string): Uint8Array {
  const hex = guid.replace(/[{}\-\s]/g, "");
  const key = new Uint8Array(16);
  for (let i = 0; i < 16; i++) {
    key[i] = parseInt(hex.substring(i * 2, i * 2 + 2), 16);
  }
  key.reverse();
  return key;
}

/** Obfuscate per spec: XOR the first 32 bytes with the key (twice through the 16 bytes). */
function obfuscate(data: Uint8Array, guid: string): Uint8Array {
  const key = odttfKey(guid);
  const out = new Uint8Array(data);
  for (let i = 0; i < Math.min(32, out.length); i++) {
    out[i] ^= key[i % 16];
  }
  return out;
}

function sampleBytes(length: number): Uint8Array {
  const bytes = new Uint8Array(length);
  for (let i = 0; i < length; i++) bytes[i] = (i * 37 + 11) % 256;
  return bytes;
}

describe("deobfuscateFont", () => {
  it("recovers the original bytes from spec-obfuscated data", () => {
    const original = sampleBytes(48);
    const obfuscated = obfuscate(original, TEST_GUID);
    expect(obfuscated).not.toEqual(original); // sanity: obfuscation changed the prefix

    const restored = deobfuscateFont(obfuscated, TEST_GUID);
    expect(restored).toEqual(original);
  });

  it("only touches the first 32 bytes", () => {
    const original = sampleBytes(64);
    const restored = deobfuscateFont(obfuscate(original, TEST_GUID), TEST_GUID);
    expect(restored.subarray(32)).toEqual(original.subarray(32));
  });

  it("does not mutate the input", () => {
    const obfuscated = obfuscate(sampleBytes(40), TEST_GUID);
    const snapshot = new Uint8Array(obfuscated);
    deobfuscateFont(obfuscated, TEST_GUID);
    expect(obfuscated).toEqual(snapshot);
  });

  it("accepts GUIDs without braces or hyphens", () => {
    const original = sampleBytes(40);
    const bare = TEST_GUID.replace(/[{}-]/g, "");
    expect(deobfuscateFont(obfuscate(original, TEST_GUID), bare)).toEqual(original);
  });

  it("returns the data unchanged for an invalid key", () => {
    const data = sampleBytes(40);
    expect(deobfuscateFont(data, "not-a-guid")).toBe(data);
    expect(deobfuscateFont(data, "")).toBe(data);
  });

  it("handles data shorter than 32 bytes", () => {
    const original = sampleBytes(10);
    const obfuscated = obfuscate(original, TEST_GUID);
    expect(deobfuscateFont(obfuscated, TEST_GUID)).toEqual(original);
  });
});

// ---------------------------------------------------------------------------
// decodeEmbeddedFont
// ---------------------------------------------------------------------------

/** Bytes that look like a raw TrueType font (sfnt version 1.0). */
function fakeTtf(length = 64): Uint8Array {
  const bytes = sampleBytes(length);
  bytes.set([0x00, 0x01, 0x00, 0x00]);
  return bytes;
}

/**
 * Build a minimal EOT container (version 0x00010000): 80-byte fixed header
 * with MagicNumber 0x504C at offset 34, four empty name strings, then the
 * font payload.
 */
function fakeEot(payload: Uint8Array): Uint8Array {
  const headerSize = 80 + 4 * 4; // fixed fields + 4 empty (pad+size) name strings
  const out = new Uint8Array(headerSize + payload.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, out.length, true); // EOTSize
  view.setUint32(4, payload.length, true); // FontDataSize
  view.setUint32(8, 0x00010000, true); // Version
  view.setUint16(34, 0x504c, true); // MagicNumber
  out.set(payload, headerSize);
  return out;
}

describe("decodeEmbeddedFont", () => {
  it("passes raw font binaries through unchanged", () => {
    const ttf = fakeTtf();
    expect(decodeEmbeddedFont(ttf)).toBe(ttf);

    const otto = sampleBytes(16);
    otto.set([0x4f, 0x54, 0x54, 0x4f]); // "OTTO"
    expect(decodeEmbeddedFont(otto)).toBe(otto);

    const woff = sampleBytes(16);
    woff.set([0x77, 0x4f, 0x46, 0x46]); // "wOFF"
    expect(decodeEmbeddedFont(woff)).toBe(woff);
  });

  it("deobfuscates ODTTF data with the font key before sniffing", () => {
    const ttf = fakeTtf(48);
    const obfuscated = obfuscate(ttf, TEST_GUID);
    expect(decodeEmbeddedFont(obfuscated, TEST_GUID)).toEqual(ttf);
  });

  it("extracts a raw font payload from an EOT container", () => {
    const payload = fakeTtf(40);
    const decoded = decodeEmbeddedFont(fakeEot(payload));
    expect(decoded).toBeDefined();
    expect(new Uint8Array(decoded!)).toEqual(payload);
  });

  it("returns undefined for empty or unrecognized data", () => {
    expect(decodeEmbeddedFont(new Uint8Array(0))).toBeUndefined();
    expect(decodeEmbeddedFont(new Uint8Array([1, 2, 3, 4, 5]))).toBeUndefined();
    // Big enough to be size-plausible but with no known magic anywhere.
    expect(decodeEmbeddedFont(new Uint8Array(100).fill(0xaa))).toBeUndefined();
  });

  it("returns undefined for an EOT whose payload is not decompressible", () => {
    // Valid EOT framing, but payload is neither a raw font nor valid MTX data.
    const junkPayload = new Uint8Array(32).fill(0x5a);
    expect(decodeEmbeddedFont(fakeEot(junkPayload))).toBeUndefined();
  });
});

describe("toStandaloneArrayBuffer", () => {
  it("copies a subarray into a standalone buffer", () => {
    const backing = sampleBytes(16);
    const slice = backing.subarray(4, 8);
    const standalone = toStandaloneArrayBuffer(slice);
    expect(standalone.byteLength).toBe(4);
    expect(new Uint8Array(standalone)).toEqual(new Uint8Array(backing.subarray(4, 8)));

    // Mutating the copy must not touch the original backing store.
    new Uint8Array(standalone)[0] = 0xff;
    expect(backing[4]).not.toBe(0xff);
  });
});
