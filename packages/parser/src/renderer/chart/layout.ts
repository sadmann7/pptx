export function numToPct(val: number): string {
  const n = Math.round(val * 10000) / 100;
  return `${Number.isInteger(n) ? n.toFixed(0) : n}%`.replace(/\.0%$/, '%');
}
