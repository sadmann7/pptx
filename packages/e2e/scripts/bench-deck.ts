/**
 * Grows a real deck to N slides for `pnpm profile:editor`:
 *
 *   pnpm bench-deck [slides] [source]
 *
 * Defaults to 300 slides of fixtures/adventure-club-pin-and-paper.pptx, written
 * to the gitignored recordings/perf/decks/. Copies drop their notes.
 */
import JSZip from "jszip";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { basename, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));
const SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

const count = Number(process.argv[2] ?? 300);
const source = process.argv[3] ?? join(HERE, "..", "fixtures", "adventure-club-pin-and-paper.pptx");
const outDir = join(HERE, "..", "recordings", "perf", "decks");

const zip = await JSZip.loadAsync(readFileSync(source));
const sourceSlides = Object.keys(zip.files)
  .filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  .sort((a, b) => Number(a.match(/\d+/)![0]) - Number(b.match(/\d+/)![0]));
const slides = await Promise.all(
  sourceSlides.map(async (path) => ({
    xml: await zip.file(path)!.async("string"),
    rels:
      (await zip.file(path.replace("slides/", "slides/_rels/") + ".rels")?.async("string")) ?? "",
  })),
);

for (const path of sourceSlides) {
  zip.remove(path);
  zip.remove(path.replace("slides/", "slides/_rels/") + ".rels");
}

const sldIds: string[] = [];
const presentationRels: string[] = [];
const overrides: string[] = [];
for (let index = 0; index < count; index++) {
  const { xml, rels } = slides[index % slides.length];
  const name = `slide${index + 1}.xml`;
  zip.file(`ppt/slides/${name}`, xml);
  zip.file(
    `ppt/slides/_rels/${name}.rels`,
    rels.replace(/<Relationship [^>]*notesSlide[^>]*\/>/g, ""),
  );
  sldIds.push(`<p:sldId id="${256 + index}" r:id="rIdBench${index}"/>`);
  presentationRels.push(
    `<Relationship Id="rIdBench${index}" Type="${SLIDE_REL}" Target="/ppt/slides/${name}"/>`,
  );
  overrides.push(`<Override PartName="/ppt/slides/${name}" ContentType="${SLIDE_TYPE}"/>`);
}

const presentation = await zip.file("ppt/presentation.xml")!.async("string");
zip.file(
  "ppt/presentation.xml",
  presentation.replace(
    /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
    `<p:sldIdLst xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${sldIds.join("")}</p:sldIdLst>`,
  ),
);
const relsPath = "ppt/_rels/presentation.xml.rels";
const rels = (await zip.file(relsPath)!.async("string")).replace(
  /<Relationship [^>]*Type="[^"]*\/slide"[^>]*\/>/g,
  "",
);
zip.file(
  relsPath,
  rels.replace("</Relationships>", `${presentationRels.join("")}</Relationships>`),
);
const types = (await zip.file("[Content_Types].xml")!.async("string")).replace(
  /<Override PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g,
  "",
);
zip.file("[Content_Types].xml", types.replace("</Types>", `${overrides.join("")}</Types>`));

mkdirSync(outDir, { recursive: true });
const out = join(outDir, `bench-${count}.pptx`);
writeFileSync(out, await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
console.log(
  `Wrote ${out}: ${count} slides from ${basename(source)} (${slides.length} source slides)`,
);
