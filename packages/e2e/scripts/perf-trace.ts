/**
 * Summarizes a trace written by `pnpm profile:editor`: where the main thread's
 * time went (scripting, style, layout, paint), the longest tasks, and the
 * functions with the most self time.
 *
 *   pnpm perf:trace recordings/perf/<label>/<scenario>.json
 *
 * Function names resolve through source maps when the docs build kept them
 * (`PROFILE=1 pnpm -F docs build`); otherwise they read as minified names.
 */
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SourceMapConsumer } from "source-map-js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const CHUNKS_DIR = path.join(HERE, "..", "..", "..", "apps", "docs", ".next", "static", "chunks");

interface TraceEvent {
  name: string;
  ph: string;
  pid: number;
  tid: number;
  ts: number;
  dur?: number;
  args?: {
    name?: string;
    data?: {
      cpuProfile?: {
        nodes?: {
          id: number;
          callFrame: {
            functionName: string;
            url: string;
            lineNumber: number;
            columnNumber: number;
          };
        }[];
        samples?: number[];
      };
      timeDeltas?: number[];
    };
  };
}

const GROUPS: Record<string, string[]> = {
  scripting: ["FunctionCall", "EvaluateScript", "TimerFire", "FireAnimationFrame", "RunMicrotasks"],
  style: ["UpdateLayoutTree"],
  layout: ["Layout"],
  paint: ["PrePaint", "Paint", "Layerize", "Commit"],
  gc: ["MinorGC", "MajorGC"],
};

const file = process.argv[2];
if (!file) throw new Error("Usage: pnpm perf:trace <trace.json>");
const { traceEvents: events } = JSON.parse(readFileSync(file, "utf8")) as {
  traceEvents: TraceEvent[];
};

const mainThreads = new Set(
  events
    .filter(
      (event) =>
        event.ph === "M" && event.name === "thread_name" && event.args?.name === "CrRendererMain",
    )
    .map((event) => `${event.pid}:${event.tid}`),
);
const busy = new Map<string, number>();
for (const event of events) {
  const thread = `${event.pid}:${event.tid}`;
  if (event.ph === "X" && event.name === "RunTask" && mainThreads.has(thread)) {
    busy.set(thread, (busy.get(thread) ?? 0) + (event.dur ?? 0));
  }
}
const pageThread = [...busy.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
const onMain = events.filter(
  (event) => `${event.pid}:${event.tid}` === pageThread && event.ph === "X",
);

const totals = Object.fromEntries(
  Object.entries(GROUPS).map(([group, names]) => [
    group,
    Math.round(
      onMain
        .filter((event) => names.includes(event.name))
        .reduce((sum, e) => sum + (e.dur ?? 0), 0) / 1000,
    ),
  ]),
);
console.log("Main thread (ms):", totals);

console.log("\nLongest tasks:");
const tasks = onMain
  .filter((event) => event.name === "RunTask")
  .sort((a, b) => (b.dur ?? 0) - (a.dur ?? 0))
  .slice(0, 5);
for (const task of tasks) {
  const end = task.ts + (task.dur ?? 0);
  const byName = new Map<string, number>();
  for (const event of onMain) {
    if (event === task || event.ts < task.ts || event.ts + (event.dur ?? 0) > end) continue;
    byName.set(event.name, (byName.get(event.name) ?? 0) + (event.dur ?? 0));
  }
  const top = [...byName.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([name, dur]) => `${name} ${Math.round(dur / 1000)}`);
  console.log(`  ${Math.round((task.dur ?? 0) / 1000)}ms  ${top.join(", ")}`);
}

const consumers = new Map<string, SourceMapConsumer | null>();
function resolve(url: string, line: number, column: number, name: string) {
  const chunk = url.split("/").pop() ?? "";
  if (!consumers.has(chunk)) {
    // Turbopack names a chunk's map separately, in its sourceMappingURL comment.
    const chunkPath = path.join(CHUNKS_DIR, chunk);
    const mapName = existsSync(chunkPath)
      ? /\/\/# sourceMappingURL=(\S+)\s*$/.exec(readFileSync(chunkPath, "utf8"))?.[1]
      : undefined;
    const mapPath = path.join(CHUNKS_DIR, mapName ?? `${chunk}.map`);
    consumers.set(
      chunk,
      existsSync(mapPath) ? new SourceMapConsumer(JSON.parse(readFileSync(mapPath, "utf8"))) : null,
    );
  }
  const consumer = consumers.get(chunk);
  if (!consumer) return `${name || "(anonymous)"} ${chunk}`;
  const original = consumer.originalPositionFor({ line: line + 1, column });
  if (!original.source) return `${name || "(anonymous)"} ${chunk}`;
  const source = original.source.replace(/^.*?(packages|apps|node_modules)\//, "$1/");
  return `${original.name ?? name ?? "(anonymous)"} ${source}:${original.line}`;
}

const nodes = new Map<number, string>();
const selfTime = new Map<string, number>();
for (const event of events) {
  if (event.name !== "ProfileChunk") continue;
  const profile = event.args?.data?.cpuProfile;
  for (const node of profile?.nodes ?? []) {
    const { functionName, url, lineNumber, columnNumber } = node.callFrame;
    nodes.set(
      node.id,
      url?.includes("/_next/")
        ? resolve(url, lineNumber, columnNumber, functionName)
        : functionName || "(native)",
    );
  }
  const deltas = event.args?.data?.timeDeltas ?? [];
  (profile?.samples ?? []).forEach((id, index) => {
    const key = nodes.get(id);
    if (!key || key === "(idle)" || key === "(program)") return;
    selfTime.set(key, (selfTime.get(key) ?? 0) + (deltas[index] ?? 0));
  });
}
console.log("\nMost self time:");
for (const [key, micros] of [...selfTime.entries()].sort((a, b) => b[1] - a[1]).slice(0, 15)) {
  console.log(`  ${String(Math.round(micros / 1000)).padStart(5)}ms  ${key}`);
}
