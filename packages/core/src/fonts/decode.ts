/**
 * Pure decode pipeline for embedded PPTX fonts (.fntdata).
 *
 * .fntdata files are EOT (Embedded OpenType) containers whose payload may be
 * MTX-compressed (MicroType Express) and/or XOR-obfuscated. This module
 * dispatches on the container's own flags and returns a raw
 * TrueType/OpenType binary.
 *
 * Kept free of DOM dependencies so it runs identically on the main thread
 * and inside a Web Worker.
 */

import { deobfuscateFont } from "./deobfuscate";
import { eotToTtf } from "./mtx";

function isRawFont(data: Uint8Array): boolean {
  if (data.length < 4) return false;
  const b0 = data[0],
    b1 = data[1],
    b2 = data[2],
    b3 = data[3];
  if (b0 === 0x00 && b1 === 0x01 && b2 === 0x00 && b3 === 0x00) return true; // TrueType
  if (b0 === 0x4f && b1 === 0x54 && b2 === 0x54 && b3 === 0x4f) return true; // OpenType CFF
  if (b0 === 0x74 && b1 === 0x72 && b2 === 0x75 && b3 === 0x65) return true; // Apple TrueType
  if (b0 === 0x74 && b1 === 0x74 && b2 === 0x63 && b3 === 0x66) return true; // TrueType collection
  if (b0 === 0x77 && b1 === 0x4f && b2 === 0x46 && b3 === 0x46) return true; // WOFF
  if (b0 === 0x77 && b1 === 0x4f && b2 === 0x46 && b3 === 0x32) return true; // WOFF2
  return false;
}

/**
 * Decode one embedded font part into a raw TrueType/OpenType binary.
 * Returns `undefined` when the data cannot be decoded.
 */
export function decodeEmbeddedFont(raw: Uint8Array, fontKey?: string): Uint8Array | undefined {
  if (raw.length === 0) return undefined;

  const data = fontKey ? deobfuscateFont(raw, fontKey) : raw;

  if (isRawFont(data)) return data;

  try {
    const decoded = eotToTtf(data);
    // An uncompressed EOT payload is passed through untouched, so the sfnt
    // signature is the only evidence that the header offsets were right.
    return isRawFont(decoded) ? decoded : undefined;
  } catch {
    return undefined;
  }
}

/** Copy a Uint8Array's contents into a standalone ArrayBuffer. */
export function toStandaloneArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}
