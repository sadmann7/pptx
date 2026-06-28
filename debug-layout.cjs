// @ts-nocheck
const JSZip = require("./node_modules/.pnpm/jszip@3.10.1/node_modules/jszip");
const {
  XMLParser,
} = require("./node_modules/.pnpm/fast-xml-parser@4.5.6/node_modules/fast-xml-parser");
const { readFileSync } = require("node:fs");

const PPTX_PATH = "[EXT] Q3 North 2025 Roadmap.pptx";

const ALWAYS_ARRAY = new Set([
  "p:sldId",
  "p:sldMasterId",
  "p:sldLayoutId",
  "p:sp",
  "p:pic",
  "p:graphicFrame",
  "p:grpSp",
  "p:cxnSp",
  "a:p",
  "a:r",
  "a:br",
  "a:fld",
  "a:tr",
  "a:tc",
  "a:gridCol",
  "a:tblStyleLst",
  "Relationship",
  "Override",
  "Default",
  "a:effectStyle",
  "a:gs",
  "a:tab",
  "c:ser",
  "c:pt",
  "c:dPt",
]);

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  allowBooleanAttributes: true,
  parseAttributeValue: false,
  parseTagValue: true,
  trimValues: true,
  isArray: (name) => ALWAYS_ARRAY.has(name),
});

function parseXml(str) {
  return parser.parse(str);
}
function attr(node, name) {
  if (!node || typeof node !== "object") return undefined;
  const val = node[`@_${name}`];
  if (val === undefined || val === null) return undefined;
  return String(val);
}
function get(root, ...path) {
  let cur = root;
  for (const key of path) {
    if (!cur || typeof cur !== "object") return undefined;
    cur = cur[key];
  }
  return cur;
}
function toArray(val) {
  if (val === undefined || val === null) return [];
  if (Array.isArray(val)) return val;
  return [val];
}
function normalizePath(p) {
  return p.replace(/\\/g, "/").replace(/^\//, "");
}
function joinPath(dir, file) {
  const parts = `${dir}/${file}`.split("/");
  const result = [];
  for (const p of parts) {
    if (p === "..") result.pop();
    else if (p !== "." && p !== "") result.push(p);
  }
  return result.join("/");
}
function toRelsPath(filePath) {
  const lastSlash = filePath.lastIndexOf("/");
  const dir = lastSlash >= 0 ? filePath.substring(0, lastSlash) : "";
  const filename = lastSlash >= 0 ? filePath.substring(lastSlash + 1) : filePath;
  return dir ? `${dir}/_rels/${filename}.rels` : `_rels/${filename}.rels`;
}
function resolveRelTarget(sourcePath, target) {
  if (target.startsWith("/")) return target.slice(1);
  const sourceDir = sourcePath.includes("/")
    ? sourcePath.substring(0, sourcePath.lastIndexOf("/"))
    : "";
  return joinPath(sourceDir, target);
}

async function loadRels(zip, filePath) {
  const relsPath = toRelsPath(filePath);
  const f = zip.file(normalizePath(relsPath));
  if (!f) {
    console.log(`  No rels file at: ${relsPath}`);
    return new Map();
  }
  const str = await f.async("string");
  const xml = parseXml(str);
  const map = new Map();
  const rels = toArray(get(xml, "Relationships", "Relationship"));
  for (const rel of rels) {
    const id = attr(rel, "Id") ?? "";
    const type = attr(rel, "Type") ?? "";
    const rawTarget = attr(rel, "Target") ?? "";
    const targetMode = attr(rel, "TargetMode");
    const target = targetMode === "External" ? rawTarget : resolveRelTarget(filePath, rawTarget);
    if (id) map.set(id, { id, type, target });
  }
  return map;
}

async function main() {
  const buf = readFileSync(PPTX_PATH);
  const zip = await JSZip.loadAsync(buf);

  const layoutPath = "ppt/slideLayouts/slideLayout15.xml";
  const layoutFile = zip.file(layoutPath);
  if (!layoutFile) {
    console.error("slideLayout15.xml NOT FOUND!");
    return;
  }

  const layoutXml = await layoutFile.async("string");
  const layoutParsed = parseXml(layoutXml);

  const cSld = get(layoutParsed, "p:sldLayout", "p:cSld");
  const spTree = get(cSld, "p:spTree");

  if (!spTree) {
    console.error("No spTree in layout!");
    return;
  }

  console.log("=== spTree keys ===");
  console.log(Object.keys(spTree));

  const picNodes = toArray(spTree["p:pic"]);
  console.log(`\n=== picNodes count: ${picNodes.length} ===`);

  for (let i = 0; i < picNodes.length; i++) {
    const pic = picNodes[i];
    const nvPicPr = get(pic, "p:nvPicPr");
    const cNvPr = get(nvPicPr, "p:cNvPr");
    const id = attr(cNvPr, "id");
    const name = attr(cNvPr, "name");

    const spPr = get(pic, "p:spPr");
    const xfrm = get(spPr, "a:xfrm");
    const off = get(xfrm, "a:off");
    const ext = get(xfrm, "a:ext");

    const blipFill = get(pic, "p:blipFill");
    const blip = get(blipFill, "a:blip");

    console.log(`\n  pic[${i}]: id=${id}, name=${name}`);
    console.log(`    position: x=${attr(off, "x")}, y=${attr(off, "y")}`);
    console.log(`    size: cx=${attr(ext, "cx")}, cy=${attr(ext, "cy")}`);
    if (blip) {
      console.log(`    blip keys: ${Object.keys(blip)}`);
      const allAttrKeys = Object.keys(blip).filter((k) => k.startsWith("@_"));
      console.log(`    @_ attrs: ${JSON.stringify(allAttrKeys)}`);
      for (const k of allAttrKeys) {
        console.log(`      ${k} = ${blip[k]}`);
      }
    } else {
      console.log(`    blip: MISSING`);
    }
  }

  // Layout rels
  const layoutRels = await loadRels(zip, layoutPath);
  console.log("\n=== Layout rels ===");
  for (const [id, rel] of layoutRels) {
    console.log(`  ${id}: ${rel.type.split("/").pop()} => ${rel.target}`);
  }

  // Check if media files exist
  console.log("\n=== Media files in zip (first 10) ===");
  const mediaFiles = Object.keys(zip.files)
    .filter((k) => k.startsWith("ppt/media/"))
    .slice(0, 10);
  mediaFiles.forEach((f) => console.log(`  ${f}`));

  // Test resolving rId2
  const picRel = layoutRels.get("rId2");
  console.log(`\n=== rId2 ===`);
  console.log(picRel ? JSON.stringify(picRel) : "NOT FOUND");
  if (picRel) {
    const mediaFile = zip.file(normalizePath(picRel.target));
    console.log(`  File in zip: ${!!mediaFile}`);
    if (mediaFile) {
      const bytes = await mediaFile.async("uint8array");
      console.log(`  Bytes length: ${bytes.length}, first bytes: ${bytes.slice(0, 4).join(",")}`);
    }
  }
}

main().catch(console.error);
