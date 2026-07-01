import { readFileSync } from "fs";

import JSZip from "./node_modules/.pnpm/jszip@3.10.1/node_modules/jszip/dist/jszip.min.js";

const buf = readFileSync(
  "c:/Users/sadman/Downloads/PPTX examples/[EXT] Q3 North 2025  Roadmap.pptx",
);
const zip = await JSZip.loadAsync(buf);

// Check master rels for theme
const masterRels = await zip.file("ppt/slideMasters/_rels/slideMaster2.xml.rels").async("string");
const themeIdx = masterRels.indexOf("relationships/theme");
const targetIdx = masterRels.indexOf("Target=", themeIdx) + 8;
const targetEnd = masterRels.indexOf('"', targetIdx);
const themeTarget = masterRels.slice(targetIdx, targetEnd);
console.log("Theme target:", themeTarget);

const resolvedPath = themeTarget.startsWith("../")
  ? "ppt/" + themeTarget.slice(3)
  : "ppt/slideMasters/" + themeTarget;
console.log("Resolved path:", resolvedPath);

const theme = await zip.file(resolvedPath).async("string");
const dk1Idx = theme.indexOf("<a:dk1>");
const dk2Idx = theme.indexOf("<a:dk2>");
const lt1Idx = theme.indexOf("<a:lt1>");
console.log("\ndk1:", theme.slice(dk1Idx, dk1Idx + 100));
console.log("dk2:", theme.slice(dk2Idx, dk2Idx + 100));
console.log("lt1:", theme.slice(lt1Idx, lt1Idx + 100));

const master = await zip.file("ppt/slideMasters/slideMaster2.xml").async("string");
const clrMapIdx = master.indexOf("<p:clrMap");
console.log("\nMaster clrMap:", master.slice(clrMapIdx, clrMapIdx + 300));

// Also print the bgFillStyleLst from theme
const bgFillIdx = theme.indexOf("<a:bgFillStyleLst>");
if (bgFillIdx >= 0) {
  console.log("\nbgFillStyleLst:", theme.slice(bgFillIdx, bgFillIdx + 500));
}

// Check slideLayout15 for its spTree (shapes that act as background)
const layout15 = await zip.file("ppt/slideLayouts/slideLayout15.xml").async("string");
const bgInLayout = layout15.indexOf("<p:bg");
console.log("\nSlideLayout15 has bg element:", bgInLayout >= 0);

// Find shapes in layout that might act as background (using sp with z-order 0)
// Check for blipFill in shapes
const blipIdx = layout15.indexOf("<a:blipFill>");
if (blipIdx >= 0) {
  console.log("\nFound blipFill in layout15 at idx:", blipIdx);
  console.log(layout15.slice(Math.max(0, blipIdx - 200), blipIdx + 400));
}
