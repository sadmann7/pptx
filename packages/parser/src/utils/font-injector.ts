/**
 * Inject embedded PPTX fonts into the DOM as @font-face rules.
 *
 * .fntdata files in PPTX are EOT (Embedded OpenType) containers whose
 * payload is typically MTX-compressed (MicroType Express). This module
 * parses the EOT header, extracts the font data, decompresses it via
 * the mtx-decompressor library, creates blob URLs, and injects @font-face
 * rules into the document head.
 *
 * Decompression is CPU-heavy (LZCOMP + adaptive Huffman per variant), so
 * fonts are processed asynchronously — one variant per event-loop turn —
 * and rules are appended incrementally. Slides render immediately with
 * fallback fonts and swap to embedded fonts as each one becomes ready.
 */

import { decompressMtx } from "mtx-decompressor";

import type {
  EmbeddedFontEntry,
  EmbeddedFontVariant,
  PresentationData,
} from "../model/presentation";
import { deobfuscateFont } from "./font-deobfuscate";

export interface FontInjectionHandle {
  dispose(): void;
}

// ── Format detection ────────────────────────────────────────────────

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

// ── EOT header parsing ─────────────────────────────────────────────

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

// ── Font resolution pipeline ────────────────────────────────────────

function resolveFontBytes(
  variant: EmbeddedFontVariant,
  fonts: Map<string, Uint8Array>,
): Uint8Array | undefined {
  const raw = fonts.get(variant.path);
  if (!raw || raw.length === 0) return undefined;

  const data = variant.fontKey ? deobfuscateFont(raw, variant.fontKey) : raw;

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

// ── @font-face CSS generation ───────────────────────────────────────

function buildFontFaceRule(family: string, weight: string, style: string, blobUrl: string): string {
  const escaped = family.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  return (
    `@font-face {\n` +
    `  font-family: "${escaped}";\n` +
    `  font-weight: ${weight};\n` +
    `  font-style: ${style};\n` +
    `  font-display: swap;\n` +
    `  src: url("${blobUrl}") format("truetype");\n` +
    `}\n`
  );
}

const VARIANTS: {
  key: keyof Pick<EmbeddedFontEntry, "regular" | "bold" | "italic" | "boldItalic">;
  weight: string;
  style: string;
}[] = [
  { key: "regular", weight: "normal", style: "normal" },
  { key: "bold", weight: "bold", style: "normal" },
  { key: "italic", weight: "normal", style: "italic" },
  { key: "boldItalic", weight: "bold", style: "italic" },
];

interface FontTask {
  typeface: string;
  weight: string;
  style: string;
  variant: EmbeddedFontVariant;
}

/** Yield to the event loop so rendering and input stay responsive. */
function nextTick(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

// ── Public API ──────────────────────────────────────────────────────

export function injectEmbeddedFonts(presentation: PresentationData): FontInjectionHandle {
  const noop: FontInjectionHandle = { dispose() {} };

  if (!presentation.embeddedFonts || presentation.embeddedFonts.length === 0) return noop;
  if (typeof document === "undefined") return noop;

  const tasks: FontTask[] = [];
  for (const entry of presentation.embeddedFonts) {
    for (const { key, weight, style } of VARIANTS) {
      const variant = entry[key];
      if (variant) tasks.push({ typeface: entry.typeface, weight, style, variant });
    }
  }
  if (tasks.length === 0) return noop;

  const blobUrls: string[] = [];
  // Same .fntdata part can back multiple typeface entries — decompress once.
  const blobUrlByPath = new Map<string, string | null>();
  let disposed = false;

  const styleEl = document.createElement("style");
  styleEl.setAttribute("data-pptx-embedded-fonts", "true");
  document.head.appendChild(styleEl);

  void (async () => {
    for (const task of tasks) {
      if (disposed) return;

      let blobUrl = blobUrlByPath.get(task.variant.path);
      if (blobUrl === undefined) {
        const bytes = resolveFontBytes(task.variant, presentation.fonts);
        if (bytes && bytes.length > 0) {
          const arrayBuffer = bytes.buffer.slice(
            bytes.byteOffset,
            bytes.byteOffset + bytes.byteLength,
          );
          const blob = new Blob([arrayBuffer as ArrayBuffer], { type: "font/ttf" });
          blobUrl = URL.createObjectURL(blob);
          blobUrls.push(blobUrl);
        } else {
          blobUrl = null;
        }
        blobUrlByPath.set(task.variant.path, blobUrl);
      }

      if (disposed) return;
      if (blobUrl) {
        styleEl.textContent += buildFontFaceRule(task.typeface, task.weight, task.style, blobUrl);
      }

      // One decompression per turn keeps the main thread responsive.
      await nextTick();
    }
  })();

  return {
    dispose() {
      disposed = true;
      styleEl.remove();
      for (const url of blobUrls) URL.revokeObjectURL(url);
      blobUrls.length = 0;
    },
  };
}
