/**
 * SafeXmlNode — adapter over fast-xml-parser's preserveOrder output.
 *
 * The reference library (aiden0z/pptx-renderer) uses a SafeXmlNode class backed
 * by browser DOMParser. We re-implement the same API on top of fast-xml-parser
 * with preserveOrder:true so the ported rendering/model code works unchanged,
 * while staying Node-compatible and keeping fast-xml-parser as the sole XML dep.
 *
 * preserveOrder output shape:
 *   [{ "a:sp": [...children...], ":@": { "@_id": "1", "@_name": "foo" } }, ...]
 *
 * Each element in the array is an object with ONE key (the tag name) pointing
 * to its children array, plus optional ":@" for attributes.
 */

import { XMLParser } from "fast-xml-parser";

// The ordered parser preserves sibling order across different tag names.
const orderedParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  allowBooleanAttributes: true,
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: true,
  preserveOrder: true,
});

/** Internal representation of one element in the preserveOrder tree. */
type ONode = Record<string, unknown>;

export class SafeXmlNode {
  private readonly _node: ONode | null;
  private readonly _tag: string | null;
  private readonly _children: ONode[] | null;

  constructor(node: ONode | null, tag?: string) {
    this._node = node;
    if (node && !tag) {
      // Determine tag from the object's first non-":@" key
      for (const k of Object.keys(node)) {
        if (k !== ":@" && k !== "#text") {
          this._tag = k;
          const v = node[k];
          this._children = Array.isArray(v) ? (v as ONode[]) : null;
          return;
        }
      }
      this._tag = null;
      this._children = null;
    } else {
      this._tag = tag ?? null;
      if (node && tag) {
        const v = node[tag];
        this._children = Array.isArray(v) ? (v as ONode[]) : null;
      } else {
        this._children = null;
      }
    }
  }

  /** Get a string attribute value, or undefined if missing. */
  attr(name: string): string | undefined {
    if (!this._node) return undefined;
    const attrs = this._node[":@"] as Record<string, unknown> | undefined;
    if (!attrs) return undefined;

    // Try exact prefixed name
    const prefixed = `@_${name}`;
    if (prefixed in attrs) {
      const v = attrs[prefixed];
      return v === undefined || v === null ? undefined : String(v);
    }

    // Handle namespace prefix: "r:id" → try "@_r:id" then scan for localName match
    const colonIdx = name.indexOf(":");
    if (colonIdx >= 0) {
      const localName = name.slice(colonIdx + 1);
      // Scan all attrs for a match on localName after any prefix
      for (const [k, v] of Object.entries(attrs)) {
        if (!k.startsWith("@_")) continue;
        const attrName = k.slice(2); // strip "@_"
        const attrLocal = attrName.includes(":") ? attrName.split(":").pop()! : attrName;
        if (attrLocal === localName && attrName.includes(":")) {
          return v === undefined || v === null ? undefined : String(v);
        }
      }
    }

    return undefined;
  }

  /** Get a numeric attribute value, or undefined if missing or not a number. */
  numAttr(name: string): number | undefined {
    const raw = this.attr(name);
    if (raw === undefined) return undefined;
    const n = Number(raw);
    return Number.isNaN(n) ? undefined : n;
  }

  /**
   * Find the first child element matching the given localName (namespace-agnostic).
   * Returns an empty SafeXmlNode if not found, so chaining never crashes.
   */
  child(localName: string): SafeXmlNode {
    if (!this._children) return new SafeXmlNode(null);
    for (const item of this._children) {
      if (!item || typeof item !== "object") continue;
      const tag = getTag(item);
      if (tag && tagLocalName(tag) === localName) {
        return new SafeXmlNode(item, tag);
      }
    }
    return new SafeXmlNode(null);
  }

  /**
   * Get child elements, optionally filtered by localName (namespace-agnostic).
   */
  children(localName?: string): SafeXmlNode[] {
    if (!this._children) return [];
    const result: SafeXmlNode[] = [];
    for (const item of this._children) {
      if (!item || typeof item !== "object") continue;
      const tag = getTag(item);
      if (!tag) continue;
      if (localName === undefined || tagLocalName(tag) === localName) {
        result.push(new SafeXmlNode(item, tag));
      }
    }
    return result;
  }

  /** Get the text content, or empty string if the element is missing. */
  text(): string {
    if (!this._children) return "";
    return collectText(this._children);
  }

  /** Whether the underlying element actually exists. */
  exists(): boolean {
    return this._node !== null && this._tag !== null;
  }

  /** All direct child elements as SafeXmlNode[]. */
  allChildren(): SafeXmlNode[] {
    return this.children();
  }

  /** The localName of the underlying element, or empty string. */
  get localName(): string {
    return this._tag ? tagLocalName(this._tag) : "";
  }

  /** Raw access to the underlying node (for rare cases needing attrs iteration). */
  get element(): ONode | null {
    return this._node;
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function getTag(obj: ONode): string | undefined {
  for (const k of Object.keys(obj)) {
    if (k !== ":@" && k !== "#text") return k;
  }
  return undefined;
}

function tagLocalName(tag: string): string {
  const idx = tag.indexOf(":");
  return idx >= 0 ? tag.slice(idx + 1) : tag;
}

function collectText(children: ONode[]): string {
  const parts: string[] = [];
  for (const item of children) {
    if (!item || typeof item !== "object") continue;
    if ("#text" in item) {
      const t = item["#text"];
      if (t !== undefined && t !== null) parts.push(String(t));
    }
    // Recurse into child elements to collect nested text
    const tag = getTag(item);
    if (tag) {
      const nested = item[tag];
      if (Array.isArray(nested)) {
        parts.push(collectText(nested as ONode[]));
      }
    }
  }
  return parts.join("");
}

// ─── Public parse function ────────────────────────────────────────────────────

/**
 * Parse an XML string into a SafeXmlNode wrapping the document element.
 */
export function parseXml(xmlString: string): SafeXmlNode {
  if (!xmlString) return new SafeXmlNode(null);
  try {
    const result = orderedParser.parse(xmlString) as ONode[];
    if (!Array.isArray(result) || result.length === 0) return new SafeXmlNode(null);
    // Skip XML declaration (<?xml ...?>) — find first real element
    let root: ONode | undefined;
    for (const item of result) {
      if (!item || typeof item !== "object") continue;
      const tag = getTag(item);
      if (tag && !tag.startsWith("?")) {
        root = item;
        break;
      }
    }
    if (!root) return new SafeXmlNode(null);
    return new SafeXmlNode(root);
  } catch (e) {
    console.warn("XML parse error:", e);
    return new SafeXmlNode(null);
  }
}
