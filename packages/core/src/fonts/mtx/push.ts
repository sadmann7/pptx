import type { Reader, Writer } from "./binary";
import { fail } from "./error";
import { read255Short } from "./varint";

// CTF hoists every value a glyph's hinting program pushes into one shared
// stream; rebuilding the glyph re-emits them as a PUSHB/PUSHW burst. Both halves
// use caller-owned buffers so a hinted glyph allocates nothing.
export function readPushValuesInto(reader: Reader, count: number, out: Int16Array): void {
  if (out.length < count) throw new RangeError("Push-value output buffer is too small");
  let length = 0;
  while (length < count) {
    const code = reader.data[reader.pos];
    if (code === 251 || code === 252) {
      if (length < 2) fail("INVALID_CTF", "Hop code has no value two positions back", reader.pos);
      reader.pos++;
      const repeated = out[length - 2]!;
      const expansion = code === 251 ? 3 : 5;
      if (length + expansion > count)
        fail("INVALID_CTF", "Hop code exceeds declared push count", reader.pos);
      out[length++] = repeated;
      out[length++] = read255Short(reader);
      out[length++] = repeated;
      if (code === 252) {
        out[length++] = read255Short(reader);
        out[length++] = repeated;
      }
    } else {
      out[length++] = read255Short(reader);
    }
  }
}

function writePushRun(
  writer: Writer,
  values: Int16Array,
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

export function writePushInstructions(writer: Writer, values: Int16Array, count: number): void {
  let start = 0;
  while (start < count) {
    const words = values[start]! < 0 || values[start]! > 255;
    let end = start + 1;
    while (end < count && end - start < 255) {
      const nextWords = values[end]! < 0 || values[end]! > 255;
      if (nextWords !== words) break;
      end++;
    }
    writePushRun(writer, values, start, end - start, words);
    start = end;
  }
}
