/**
 * Embedded font support (separate entry point).
 *
 * Kept out of the main package entry so the decode pipeline (EOT parsing,
 * MTX decompression, worker pool) is only loaded by consumers that actually
 * inject embedded fonts, e.g. via a dynamic `import("@diceui/pptx-core/fonts")`.
 */

export { decodeEmbeddedFont, decodeEmbeddedFontPart } from "./decode";
export type { FontDecodeResult } from "./decode";
export { deobfuscateFont } from "./deobfuscate";
export { findPriorityTypefaces, loadEmbeddedFonts } from "./loader";
export type { EmbeddedFontError, EmbeddedFontsHandle, LoadEmbeddedFontsOptions } from "./loader";
export type { MtxErrorCode } from "./mtx";
