/**
 * Minimal `--flag value` parser for the drivers in this folder.
 *
 * Flags declared in `lists` may repeat and collect; flags in `numbers` are
 * coerced; anything else is a string. Bare `--flag` (no value) is `true`.
 */
export function parseArgs(argv, { numbers = [], lists = [], strings = [] } = {}) {
  const known = new Set([...numbers, ...lists, ...strings]);
  const result = {};
  for (let i = 0; i < argv.length; i++) {
    const token = argv[i];
    if (!token.startsWith("--")) {
      (result._ ??= []).push(token);
      continue;
    }
    const name = token.slice(2);
    if (!known.has(name)) throw new Error(`unknown flag --${name}`);
    const next = argv[i + 1];
    const value = next === undefined || next.startsWith("--") ? true : argv[++i];
    if (lists.includes(name)) (result[name] ??= []).push(value);
    else if (numbers.includes(name)) result[name] = Number(value);
    else result[name] = value;
  }
  return result;
}

/** Parse `start:end` (inclusive) or a single number into a range. */
export function parseRange(value) {
  const [start, end] = String(value).split(":");
  const from = Number(start);
  return { from, to: end === undefined ? from : Number(end) };
}
