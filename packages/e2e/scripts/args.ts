/** Minimal `--flag value` parsing for the debugging scripts. */

export interface ArgSpec {
  numbers?: readonly string[];
  strings?: readonly string[];
  lists?: readonly string[];
}

export interface ParsedArgs {
  positional: string[];
  numbers: Record<string, number | undefined>;
  strings: Record<string, string | undefined>;
  lists: Record<string, string[]>;
}

export function parseArgs(argv: readonly string[], spec: ArgSpec): ParsedArgs {
  const parsed: ParsedArgs = { positional: [], numbers: {}, strings: {}, lists: {} };
  for (const name of spec.lists ?? []) parsed.lists[name] = [];

  for (let i = 0; i < argv.length; i++) {
    const token = argv[i];
    if (!token.startsWith("--")) {
      parsed.positional.push(token);
      continue;
    }
    const name = token.slice(2);
    const value = argv[++i];
    if (value === undefined) throw new Error(`--${name} needs a value`);

    if (spec.numbers?.includes(name)) {
      const parsedNumber = Number(value);
      if (Number.isNaN(parsedNumber)) throw new Error(`--${name} expects a number, got "${value}"`);
      parsed.numbers[name] = parsedNumber;
    } else if (spec.lists?.includes(name)) {
      parsed.lists[name].push(value);
    } else if (spec.strings?.includes(name)) {
      parsed.strings[name] = value;
    } else {
      throw new Error(`unknown option --${name}`);
    }
  }

  return parsed;
}

/** Parses a "from:to" range, defaulting either end to the given bounds. */
export function parseRange(value: string | undefined, from: number, to: number): [number, number] {
  if (!value) return [from, to];
  const [start, end] = value.split(":");
  return [start ? Number(start) : from, end ? Number(end) : to];
}
