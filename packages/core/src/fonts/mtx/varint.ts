/**
 * CTF variable-length integers (MTX specification, section 6.1).
 *
 * Both types spend one byte on values in the common range and escape through
 * codes 250 and 253-255 for anything wider.
 */

import type { Reader } from "./binary";
import { fail } from "./error";

export function read255UShort(reader: Reader): number {
  const code = reader.u8();
  if (code === 253) return reader.u16be();
  if (code === 254) return 506 + reader.u8();
  if (code === 255) return 253 + reader.u8();
  return code;
}

export function read255Short(reader: Reader): number {
  let code = reader.u8();
  if (code === 253) return reader.i16be();
  let sign = 1;
  if (code === 250) {
    sign = -1;
    code = reader.u8();
  }
  if (code === 251 || code === 252 || code === 250 || code === 253) {
    fail("INVALID_CTF", "Reserved code inside a 255SHORT value", reader.pos - 1);
  }
  const magnitude = code === 255 ? 250 + reader.u8() : code === 254 ? 500 + reader.u8() : code;
  return magnitude * sign;
}
