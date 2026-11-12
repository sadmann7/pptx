// Pulls the Presentation components from the diceui registry, which owns them.
//
//   pnpm ui:sync    replace the local copies with the registry's.
//   pnpm ui:check   fail if the local copies differ from the registry's.
import { spawnSync } from "node:child_process";
import { readdir, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";

const ROOT = path.resolve(import.meta.dirname, "..");
const ITEMS = ["@diceui/presentation", "@diceui/presentation-zoom-select"];
const FILES = [
  "src/components/presentation.tsx",
  "src/components/presentation-zoom-select.tsx",
];

const isCheck = process.argv.includes("--check");
const snapshot = await readTree("src");

// The items' own dependencies (button, select, use-mobile) are pptx's to keep,
// so everything the add touches besides FILES is put back afterwards.
const ok =
  run("pnpm", [
    "dlx",
    "shadcn@latest",
    "add",
    ...ITEMS,
    "--yes",
    "--overwrite",
  ]) && run("pnpm", ["exec", "oxfmt", ...FILES]);
const synced = await readTree("src");
await restore((file) => !FILES.includes(file) || !ok || isCheck);
if (!ok) process.exit(1);

const changed = FILES.filter((file) => snapshot.get(file) !== synced.get(file));
if (isCheck && changed.length > 0) {
  console.error(
    `Out of sync with the diceui registry: ${changed.join(", ")}. Run pnpm ui:sync.`,
  );
  process.exit(1);
}
console.log(
  changed.length > 0 && !isCheck
    ? `Updated ${changed.join(", ")}.`
    : "In sync.",
);

function run(command: string, args: string[]) {
  const result = spawnSync(command, args, { cwd: ROOT, encoding: "utf8" });
  if (result.status === 0) return true;
  console.error(result.stdout, result.stderr);
  return false;
}

async function readTree(dir: string) {
  const entries = await readdir(path.join(ROOT, dir), {
    recursive: true,
    withFileTypes: true,
  });
  const files = new Map<string, string>();
  for (const entry of entries) {
    if (!entry.isFile()) continue;
    const file = path.relative(ROOT, path.join(entry.parentPath, entry.name));
    files.set(file, await readFile(path.join(ROOT, file), "utf8"));
  }
  return files;
}

async function restore(shouldRestore: (file: string) => boolean) {
  const current = await readTree("src");
  for (const file of new Set([...snapshot.keys(), ...current.keys()])) {
    if (!shouldRestore(file) || snapshot.get(file) === current.get(file))
      continue;
    const source = snapshot.get(file);
    const target = path.join(ROOT, file);
    if (source === undefined) await rm(target, { force: true });
    else await writeFile(target, source);
  }
}
