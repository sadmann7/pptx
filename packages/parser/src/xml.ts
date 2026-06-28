import { XMLParser } from "fast-xml-parser";

/**
 * Elements that must always be treated as arrays even when there's only
 * one instance in the XML.
 */
const ALWAYS_ARRAY = new Set([
  // Presentation — only the repeating ID elements, NOT their container wrappers
  "p:sldId",
  "p:sldMasterId",
  "p:sldLayoutId",
  // Slide content — repeating shape elements only; p:spTree is a single container
  "p:sp",
  "p:pic",
  "p:graphicFrame",
  "p:grpSp",
  "p:cxnSp",
  // Text
  "a:p",
  "a:r",
  "a:br",
  "a:fld",
  // Table
  "a:tr",
  "a:tc",
  "a:gridCol",
  "a:tblStyleLst",
  // Relationships
  "Relationship",
  // Content types
  "Override",
  "Default",
  // Theme — effectStyle repeats within effectStyleLst
  "a:effectStyle",
  // Gradient stops
  "a:gs",
  // Paragraph spacing
  "a:tab",
  // DrawingML Charts — repeating elements
  "c:ser", // chart series
  "c:pt", // data points in strCache / numCache
  "c:dPt", // per-point style overrides
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

export function parseXml(xmlString: string): Record<string, unknown> {
  return parser.parse(xmlString) as Record<string, unknown>;
}

/** Safe attribute getter — returns string or undefined */
export function attr(node: unknown, name: string): string | undefined {
  if (node === null || typeof node !== "object") return undefined;
  const val = (node as Record<string, unknown>)[`@_${name}`];
  if (val === undefined || val === null) return undefined;
  return String(val);
}

/** Safe numeric attribute getter */
export function attrNum(node: unknown, name: string): number | undefined {
  const s = attr(node, name);
  if (s === undefined) return undefined;
  const n = Number(s);
  return Number.isNaN(n) ? undefined : n;
}

/** Safe boolean attribute getter — treats '1', 'true' as true */
export function attrBool(node: unknown, name: string): boolean | undefined {
  const s = attr(node, name);
  if (s === undefined) return undefined;
  return s === "1" || s === "true";
}

/**
 * Get a child node by dotted path, e.g. get(root, 'p:sld', 'p:cSld', 'p:spTree')
 * Returns undefined if any segment is missing.
 */
export function get(root: unknown, ...path: string[]): unknown {
  let cur: unknown = root;
  for (const key of path) {
    if (cur === null || typeof cur !== "object") return undefined;
    cur = (cur as Record<string, unknown>)[key];
  }
  return cur;
}

/**
 * Returns a node as an array. If it's already an array, return it.
 * If it's a non-null value, wrap it. If undefined/null, return [].
 */
export function toArray<T>(val: T | T[] | undefined | null): T[] {
  if (val === undefined || val === null) return [];
  if (Array.isArray(val)) return val;
  return [val];
}

/** Get text content of a node (handles both string leaf and object with #text) */
export function textContent(node: unknown): string {
  if (typeof node === "string") return node;
  if (typeof node === "number") return String(node);
  if (node === null || node === undefined) return "";
  if (typeof node === "object") {
    const t = (node as Record<string, unknown>)["#text"];
    if (t !== undefined) return String(t);
  }
  return "";
}

// ─── Document-order extraction ────────────────────────────────────────────────

const SHAPE_CHILD_TAGS = new Set(["p:sp", "p:pic", "p:graphicFrame", "p:cxnSp", "p:grpSp"]);

export interface ChildOrderNode {
  tag: string;
  groupChildren?: ChildOrderNode[];
}

const orderedParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  allowBooleanAttributes: true,
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: true,
  preserveOrder: true,
});

/**
 * Extract the document-order child tag sequence of a spTree node.
 * Uses a separate parser with `preserveOrder: true` so cross-type
 * interleaving (p:sp, p:pic, p:grpSp, …) is preserved.
 */
export function extractSpTreeChildOrder(
  xmlString: string,
  pathToSpTree: string[],
): ChildOrderNode[] {
  const ordered = orderedParser.parse(xmlString) as unknown[];
  return navigateOrderedTree(ordered, pathToSpTree);
}

function navigateOrderedTree(tree: unknown[], path: string[]): ChildOrderNode[] {
  let current: unknown[] = tree;
  for (const tag of path) {
    const found = (current as Record<string, unknown>[])?.find(
      (item) => item != null && typeof item === "object" && tag in item,
    );
    if (!found) return [];
    const next = (found as Record<string, unknown>)[tag];
    if (!Array.isArray(next)) return [];
    current = next;
  }
  return extractOrderFromChildren(current);
}

function extractOrderFromChildren(children: unknown[]): ChildOrderNode[] {
  const result: ChildOrderNode[] = [];
  for (const item of children) {
    if (!item || typeof item !== "object") continue;
    const keys = Object.keys(item as object).filter((k) => k !== ":@");
    const tag = keys[0];
    if (!tag || !SHAPE_CHILD_TAGS.has(tag)) continue;
    const node: ChildOrderNode = { tag };
    if (tag === "p:grpSp") {
      const grpChildren = (item as Record<string, unknown>)[tag];
      if (Array.isArray(grpChildren)) {
        node.groupChildren = extractOrderFromChildren(grpChildren);
      }
    }
    result.push(node);
  }
  return result;
}
