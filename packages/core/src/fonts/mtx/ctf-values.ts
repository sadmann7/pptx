import { Reader, Writer } from "./binary";
import { fail } from "./errors";

export function read255UShort(reader: Reader): number {
  const code = reader.u8();
  if (code === 253) return reader.u16be();
  if (code === 254) return 506 + reader.u8();
  if (code === 255) return 253 + reader.u8();
  return code;
}

function read255ShortValue(reader: Reader): number {
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
          ? [repeated, read255ShortValue(reader), repeated]
          : [repeated, read255ShortValue(reader), repeated, read255ShortValue(reader), repeated];
      if (out.length + expansion.length > count)
        fail("INVALID_CTF", "Hop code exceeds declared push count", reader.pos);
      out.push(...expansion);
    } else {
      out.push(read255ShortValue(reader));
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

export function decodeCvt(encoded: Uint8Array): Uint8Array {
  const reader = new Reader(encoded);
  const count = reader.u16be();
  const writer = new Writer(count * 2);
  let previous = 0;
  for (let i = 0; i < count; i++) {
    const code = reader.u8();
    let delta: number;
    if (code < 238) delta = code;
    else if (code === 238) delta = reader.i16be();
    else if (code < 248) delta = -(238 * (code - 239) + reader.u8());
    else delta = 238 * (code - 247) + reader.u8();
    previous = (previous + delta) & 0xffff;
    writer.u16be(previous);
  }
  return writer.finish();
}
