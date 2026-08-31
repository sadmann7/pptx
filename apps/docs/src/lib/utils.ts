export function parseArrayParam<T extends string>(
  value: string | string[] | undefined,
  allowed: ReadonlySet<T>,
): ReadonlySet<T> {
  const raw = Array.isArray(value) ? value : (value?.replace(/^\[|\]$/g, "").split(",") ?? []);
  return new Set(raw.filter((v): v is T => allowed.has(v as T)));
}
