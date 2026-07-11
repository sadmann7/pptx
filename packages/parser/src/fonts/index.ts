/**
 * Embedded font support (separate entry point).
 *
 * Kept out of the main package entry so the decode pipeline (EOT parsing,
 * MTX decompression, worker pool) is only loaded by consumers that actually
 * inject embedded fonts, e.g. via a dynamic `import("@diceui/pptx-parser/fonts")`.
 */

export { decodeEmbeddedFont } from "./decode";
export { deobfuscateFont } from "./deobfuscate";
export { collectPriorityTypefaces, injectEmbeddedFonts } from "./injector";
export type { FontInjectionHandle, InjectEmbeddedFontsOptions } from "./injector";
