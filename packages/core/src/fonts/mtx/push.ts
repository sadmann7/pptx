/**
 * The CTF glyph push stream and its TrueType PUSH counterpart.
 *
 * CTF hoists every value a glyph's hinting program pushes into one shared
 * stream, so reconstructing the glyph means reading those values back and
 * re-emitting them as a PUSHB/PUSHW burst ahead of the bytecode.
 */

import { type Reader, Writer } from "./binary";
import { fail } from "./error";
import { read255Short } from "./varint";

export function readPushValues(reader: Reader, count: number): number[] {
  const out: number[] = [];
  while (out.length < count) {
    const code = reader.data[reader.pos];
    if (code === 251 || code === 252) {
      if (out.length < 2)
        fail("INVALID_CTF", "Hop code has no value two positions back", reader.pos);
      reader.pos++;
      const repeated = out[out.length - 2]!;
      const expansion =
        code === 251
          ? [repeated, read255Short(reader), repeated]
          : [repeated, read255Short(reader), repeated, read255Short(reader), repeated];
      if (out.length + expansion.length > count)
        fail("INVALID_CTF", "Hop code exceeds declared push count", reader.pos);
      out.push(...expansion);
    } else {
      out.push(read255Short(reader));
    }
  }
  return out;
}

function writePushRun(
  writer: Writer,
  values: number[],
  start: number,
  count: number,
  words: boolean,
): void {
  if (count <= 8) writer.u8((words ? 0xb8 : 0xb0) + count - 1);
  else {
    writer.u8(words ? 0x41 : 0x40);
    writer.u8(count);
  }
  for (let i = start; i < start + count; i++) {
    if (words) writer.i16be(values[i]!);
    else writer.u8(values[i]!);
  }
}

/** Emit a compact, semantically equivalent TrueType PUSH instruction burst. */
export function encodePushInstructions(values: number[]): Uint8Array {
  const writer = new Writer(values.length * 2 + 8);
  let start = 0;
  while (start < values.length) {
    const words = values[start]! < 0 || values[start]! > 255;
    let end = start + 1;
    while (end < values.length && end - start < 255) {
      const nextWords = values[end]! < 0 || values[end]! > 255;
      if (nextWords !== words) break;
      end++;
    }
    writePushRun(writer, values, start, end - start, words);
    start = end;
  }
  return writer.finish();
}
