/**
 * Pure decode pipeline for embedded PPTX fonts (.fntdata).
 *
 * .fntdata files are EOT (Embedded OpenType) containers whose payload is
 * typically MTX-compressed (MicroType Express). This module parses the EOT
 * header, extracts the font data, and decompresses it into a raw
 * TrueType/OpenType binary.
 *
 * Kept free of DOM dependencies so it runs identically on the main thread
 * and inside a Web Worker.
 */

import { decompressMtx } from "mtx-decompressor";

import { deobfuscateFont } from "./font-deobfuscate";

function isRawFont(data: Uint8Array): boolean {
  if (data.length < 4) return false;
  const b0 = data[0],
    b1 = data[1],
    b2 = data[2],
    b3 = data[3];
  if (b0 === 0x00 && b1 === 0x01 && b2 === 0x00 && b3 === 0x00) return true; // TrueType
  if (b0 === 0x4f && b1 === 0x54 && b2 === 0x54 && b3 === 0x4f) return true; // OpenType CFF
  if (b0 === 0x74 && b1 === 0x72 && b2 === 0x75 && b3 === 0x65) return true; // Apple TrueType
  if (b0 === 0x77 && b1 === 0x4f && b2 === 0x46 && b3 === 0x46) return true; // WOFF
  if (b0 === 0x77 && b1 === 0x4f && b2 === 0x46 && b3 === 0x32) return true; // WOFF2
  return false;
}

function isEot(data: Uint8Array): boolean {
  if (data.length < 36) return false;
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  return view.getUint16(34, true) === 0x504c; // MagicNumber
}

/**
 * Parse the EOT container header and return the FontData payload.
 *
 * Header (little-endian):
 *   0: uint32 EOTSize
 *   4: uint32 FontDataSize
 *   8: uint32 Version
 *  12: uint32 Flags
 *  16-79: fixed fields (PANOSE, charset, weight, fsType, magic, ranges...)
 *  80+: 4 variable-length name strings, each: uint16 pad + uint16 size + bytes
 *  then version-dependent blocks (RootString, signature, EUDC)
 *  then: FontData[FontDataSize]
 */
function extractEotPayload(data: Uint8Array): Uint8Array | undefined {
  if (data.length < 82) return undefined;
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);

  const fontDataSize = view.getUint32(4, true);
  const version = view.getUint32(8, true);

  let offset = 80;
  for (let i = 0; i < 4; i++) {
    if (offset + 4 > data.length) return undefined;
    const strSize = view.getUint16(offset + 2, true);
    offset += 4 + strSize;
  }

  if (version >= 0x00020001 && offset + 4 <= data.length) {
    const rootStringSize = view.getUint16(offset + 2, true);
    offset += 4 + rootStringSize;
  }

  if (version >= 0x00020002 && offset + 8 <= data.length) {
    offset += 4; // RootStringCheckSum
    offset += 4; // EUDCCodePage
    if (offset + 4 <= data.length) {
      const signatureSize = view.getUint16(offset + 2, true);
      offset += 4 + signatureSize;
    }
    if (offset + 8 <= data.length) {
      const eudcFontSize = view.getUint32(offset + 4, true);
      offset += 8 + eudcFontSize;
    }
  }

  if (offset >= data.length) return undefined;
  const end = Math.min(offset + fontDataSize, data.length);
  return data.subarray(offset, end);
}

/**
 * Decode one embedded font part into a raw TrueType/OpenType binary.
 * Returns `undefined` when the data cannot be decoded.
 */
export function decodeEmbeddedFont(raw: Uint8Array, fontKey?: string): Uint8Array | undefined {
  if (raw.length === 0) return undefined;

  const data = fontKey ? deobfuscateFont(raw, fontKey) : raw;

  if (isRawFont(data)) return data;

  if (isEot(data)) {
    const payload = extractEotPayload(data);
    if (!payload) return undefined;
    if (isRawFont(payload)) return payload;

    try {
      return decompressMtx(payload, { compressed: true, encrypted: false });
    } catch {
      try {
        return decompressMtx(payload, { compressed: true, encrypted: true });
      } catch {
        return undefined;
      }
    }
  }

  return undefined;
}

/** Copy a Uint8Array's contents into a standalone ArrayBuffer. */
export function toStandaloneArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}
