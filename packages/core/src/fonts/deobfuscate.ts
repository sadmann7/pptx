/**
 * Undoes ECMA-376 Part 4 §17.8.1 obfuscation: the first 32 bytes are XORed with
 * the `fontKey` GUID's 16 bytes in reverse order. Returns a copy.
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
