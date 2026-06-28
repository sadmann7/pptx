import JSZip from "./node_modules/.pnpm/jszip@3.10.1/node_modules/jszip/dist/jszip.min.js";
import { readFileSync } from "fs";

const buf = readFileSync(
  "c:/Users/sadman/Downloads/PPTX examples/[EXT] Q3 North 2025  Roadmap.pptx",
);
const zip = await JSZip.loadAsync(buf);

// Check slideLayout15 spTree for large background shapes
const layout15 = await zip.file("ppt/slideLayouts/slideLayout15.xml").async("string");
console.log("Layout15 length:", layout15.length);
console.log("\n=== First 3000 chars of layout15 ===");
console.log(layout15.slice(0, 3000));

// Check layout rels
const layoutRels = await zip.file("ppt/slideLayouts/_rels/slideLayout15.xml.rels").async("string");
console.log("\n=== Layout15 rels ===");
console.log(layoutRels);

// Check slide1 - does it have any shapes that look like background images?
const slide1 = await zip.file("ppt/slides/slide1.xml").async("string");
console.log("\n=== Slide1 spTree (first 3000 chars) ===");
const spTreeIdx = slide1.indexOf("<p:spTree>");
console.log(slide1.slice(spTreeIdx, spTreeIdx + 3000));

// Check slide1 rels
const slide1Rels = await zip.file("ppt/slides/_rels/slide1.xml.rels").async("string");
console.log("\n=== Slide1 rels ===");
console.log(slide1Rels);
