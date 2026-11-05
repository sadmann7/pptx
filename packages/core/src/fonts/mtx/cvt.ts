import { Reader, Writer } from "./binary";

// MTX spec §5.2: each control value is a delta from the previous one, packed
// into one to three bytes because neighbouring values are usually close.
export function decodeCvt(encoded: Uint8Array): Uint8Array {
  const reader = new Reader(encoded);
  // A count of entries, not a byte length: the table is twice this long.
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
