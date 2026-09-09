import { decodeEmbeddedFont } from "@diceui/pptx-core/fonts";
import JSZip from "jszip";
/**
 * Classifies the embedded font parts of every deck in fixtures/ and decks/, or
 * in directories given as arguments.
 *
 * What matters for the decode cost is whether a part is MTX-compressed: an
 * uncompressed EOT payload is passed through, so it costs nothing, while MTX
 * means LZCOMP over three streams, which is the only expensive case. Which one a
 * deck carries follows from what wrote it, so the producer is reported too.
 */
import fs from "node:fs";
import path from "node:path";

import {
  parseEotMetadata,
  TTEMBED_SUBSET,
  TTEMBED_TTCOMPRESSED,
  TTEMBED_XORENCRYPTDATA,
} from "../../core/src/fonts/mtx/index.js";

const kb = (bytes: number) => `${(bytes / 1024).toFixed(0)} KB`;

/** Whatever wrote the file says so in docProps/app.xml. */
async function readProducer(zip: JSZip): Promise<string> {
  const app = await zip.file("docProps/app.xml")?.async("string");
  if (!app) return "unknown";
  const name = /<Application>([^<]*)<\/Application>/.exec(app)?.[1] ?? "unknown";
  const version = /<AppVersion>([^<]*)<\/AppVersion>/.exec(app)?.[1];
  return version ? `${name} ${version}` : name;
}

/** Which parts the deck declares a fontKey for, i.e. which are XOR-obfuscated. */
async function readObfuscatedParts(zip: JSZip): Promise<number> {
  const presentation = await zip.file("ppt/presentation.xml")?.async("string");
  if (!presentation) return 0;
  return (presentation.match(/fontKey=/g) ?? []).length;
}

/** Reports one part the same way a deck's parts are reported. */
function describePart(label: string, bytes: Uint8Array): string {
  let flags = "unknown";
  try {
    const metadata = parseEotMetadata(bytes);
    flags = [
      `EOT v${metadata.version}`,
      metadata.flags & TTEMBED_TTCOMPRESSED ? "MTX" : "plain",
      metadata.flags & TTEMBED_SUBSET ? "subset" : "full",
      metadata.flags & TTEMBED_XORENCRYPTDATA ? "xor" : "",
    ]
      .filter(Boolean)
      .join(" ");
  } catch (error) {
    flags = `header error (${error instanceof Error ? error.message : String(error)})`;
  }

  const start = performance.now();
  const decoded = decodeEmbeddedFont(bytes);
  const ms = performance.now() - start;

  return `    ${label.padEnd(40)}${kb(bytes.length).padEnd(9)}${flags.padEnd(28)}-> ${(decoded ? kb(decoded.length) : "failed").padEnd(9)}${ms.toFixed(2)}ms`;
}

// The loose parts the core decoder is tested against, for comparison: they came
// out of PowerPoint, which is the only producer here that compresses.
const CORE_FIXTURES = path.join(
  import.meta.dirname,
  "..",
  "..",
  "core",
  "src",
  "tests",
  "fixtures",
);
if (fs.existsSync(CORE_FIXTURES)) {
  console.log("core decode fixtures");
  for (const name of fs.readdirSync(CORE_FIXTURES).filter((n) => n.endsWith(".fntdata"))) {
    console.log(
      describePart(name, new Uint8Array(fs.readFileSync(path.join(CORE_FIXTURES, name)))),
    );
  }
}

// Directories to scan: the package's own by default, or whatever was asked for,
// which is how a downloaded corpus gets classified without being copied in.
const dirs = process.argv.slice(2);
for (const dir of dirs.length > 0 ? dirs : ["fixtures", "decks"]) {
  const abs = path.resolve(import.meta.dirname, "..", dir);
  if (!fs.existsSync(abs)) continue;

  for (const name of fs.readdirSync(abs).filter((n) => n.endsWith(".pptx"))) {
    // A scanned corpus can hold things that are not readable packages at all
    // (truncated downloads, LFS pointers); they are not what this reports on.
    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(fs.readFileSync(path.join(abs, name)));
    } catch {
      console.log(`${dir}/${name}\n    not a readable package, skipped`);
      continue;
    }
    const parts = Object.values(zip.files).filter((f) => /\.fntdata$/i.test(f.name));
    if (parts.length === 0) continue;

    const rows: string[] = [];
    for (const part of parts) {
      const bytes = await part.async("uint8array");
      if (bytes.length === 0) continue;
      rows.push(describePart(path.basename(part.name), bytes));
    }
    console.log(
      `${dir}/${name}\n    ${await readProducer(zip)}, ${await readObfuscatedParts(zip)} fontKey references`,
    );
    console.log(rows.join("\n"));
  }
}
