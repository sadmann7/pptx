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
import { eotToTtf, MtxError, type MtxErrorCode, parseEotMetadata } from "./mtx";

function getIsRawFont(data: Uint8Array): boolean {
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
 * The outcome of decoding one part, carrying why it failed when it did, so
 * callers can report a font that silently fell back.
 */
export type FontDecodeResult =
  | { ok: true; bytes: Uint8Array }
  | { ok: false; message: string; code?: MtxErrorCode };

/**
 * Decode one embedded font part into a raw TrueType/OpenType binary, keeping
 * the reason on failure.
 */
export function decodeEmbeddedFontPart(part: Uint8Array, fontKey?: string): FontDecodeResult {
  if (part.length === 0) return { ok: false, message: "Font part is empty" };

  const data = fontKey ? deobfuscateFont(part, fontKey) : part;

  if (getIsRawFont(data)) return { ok: true, bytes: data };

  try {
    const decoded = eotToTtf(data);
    // An uncompressed EOT payload is passed through untouched, so the sfnt
    // signature is the only evidence that the header offsets were right.
    if (getIsRawFont(decoded)) return { ok: true, bytes: decoded };
    return { ok: false, message: "Decoded payload does not start with an sfnt signature" };
  } catch (error) {
    if (error instanceof MtxError) return { ok: false, message: error.message, code: error.code };
    return { ok: false, message: error instanceof Error ? error.message : String(error) };
  }
}

/**
 * Decode one embedded font part into a raw TrueType/OpenType binary.
 * Returns `undefined` when the data cannot be decoded.
 */
export function decodeEmbeddedFont(part: Uint8Array, fontKey?: string): Uint8Array | undefined {
  const result = decodeEmbeddedFontPart(part, fontKey);
  return result.ok ? result.bytes : undefined;
}

/**
 * Whether decoding this part will run the MTX decompressor.
 *
 * That is the only expensive case: an uncompressed EOT payload is passed
 * through, so decoding it is a header read and costs microseconds, while MTX
 * means LZCOMP over three streams (a few ms per part). Callers use this to
 * decide whether the work is worth moving off the main thread at all.
 *
 * Reading the flag repeats the deobfuscation `decodeEmbeddedFont` will do
 * again, so an obfuscated part is copied here to XOR its first 32 bytes back.
 * A memcpy stays orders of magnitude below the decompression it decides on.
 */
export function getIsCompressedFont(part: Uint8Array, fontKey?: string): boolean {
  if (part.length === 0) return false;
  const data = fontKey ? deobfuscateFont(part, fontKey) : part;
  if (getIsRawFont(data)) return false;
  try {
    return parseEotMetadata(data).compressed;
  } catch {
    // Unreadable header: decoding will fail rather than decompress.
    return false;
  }
}

/**
 * Copy the bytes into an ArrayBuffer of their own, detached from any larger
 * buffer they were a view into, so it can be transferred to a worker.
 */
export function copyToArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}
