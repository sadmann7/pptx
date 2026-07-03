/**
 * ODTTF / Embedded OpenType font deobfuscation.
 *
 * ECMA-376 Part 4 §17.8.1: embedded fonts are obfuscated by XOR-ing
 * the first 32 bytes with a key derived from a GUID stored in the
 * `fontKey` attribute. The algorithm:
 *
 * 1. Strip braces/hyphens from the GUID to get 32 hex chars (16 bytes).
 * 2. Reverse the byte order (big-endian).
 * 3. XOR bytes 0–15 of the font with the 16-byte key.
 * 4. XOR bytes 16–31 of the font with the same 16-byte key.
 *
 * The rest of the file is unmodified.
 */

/**
 * Deobfuscate an ODTTF/fntdata font using the GUID-based XOR algorithm.
 * Returns a new Uint8Array with the first 32 bytes restored; the input
 * is not mutated.
 */
export function deobfuscateFont(data: Uint8Array, fontKey: string): Uint8Array {
  const hex = fontKey.replace(/[{}\-\s]/g, "");
  if (hex.length !== 32) return data;

  const keyBytes = new Uint8Array(16);
  for (let i = 0; i < 16; i++) {
    keyBytes[i] = parseInt(hex.substring(i * 2, i * 2 + 2), 16);
  }
  keyBytes.reverse();

  const result = new Uint8Array(data.length);
  result.set(data);
  const xorLen = Math.min(32, result.length);
  for (let i = 0; i < xorLen; i++) {
    result[i] ^= keyBytes[i % 16];
  }
  return result;
}
