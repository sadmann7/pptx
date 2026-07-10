/**
 * Safe XML parser using browser DOMParser.
 * All operations are null-safe — accessing missing elements never crashes.
 */

/**
 * Lazily built per-element index of children by localName.
 *
 * Parsed OOXML documents are never mutated after parsing, so the index cannot
 * go stale. Renderers query the same elements repeatedly (e.g. a shape's
 * `spPr` is probed for a dozen different child names), which made repeated
 * linear scans a measurable share of parse/render profiles.
 *
 * Two-phase to avoid taxing the initial single-pass model build (where most
 * elements are queried exactly once): the first lookup on an element is a
 * plain scan; the index is built only when the same element is queried again.
 */
const childIndexCache = new WeakMap<Element, Map<string, Element[]>>();
const queriedOnce = new WeakSet<Element>();

/**
 * Drop the cached child index for an element whose children were mutated.
 *
 * Edit operations that insert or remove child elements MUST call this on the
 * parent, otherwise later `child()`/`children()` lookups return stale results.
 * Attribute and text mutations don't affect the index and need no
 * invalidation.
 */
export function invalidateChildIndex(el: Element): void {
  childIndexCache.delete(el);
  queriedOnce.delete(el);
}

function getChildIndex(el: Element): Map<string, Element[]> | undefined {
  let index = childIndexCache.get(el);
  if (index !== undefined) return index;

  if (!queriedOnce.has(el)) {
    queriedOnce.add(el);
    return undefined;
  }

  index = new Map();
  const children = el.children;
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    const bucket = index.get(child.localName);
    if (bucket === undefined) {
      index.set(child.localName, [child]);
    } else {
      bucket.push(child);
    }
  }
  childIndexCache.set(el, index);
  return index;
}

export class SafeXmlNode {
  private readonly el: Element | null;

  constructor(el: Element | null) {
    this.el = el;
  }

  /**
   * Shared immutable instance for all "missing element" results.
   * SafeXmlNode has no mutable state, so every miss can safely alias one
   * object — child-chain misses are extremely common (most optional OOXML
   * elements are absent), and per-miss allocations dominated GC pressure
   * in parse profiles.
   */
  static readonly EMPTY = new SafeXmlNode(null);

  /** Get a string attribute value, or undefined if missing. */
  attr(name: string): string | undefined {
    if (!this.el) return undefined;
    // Fast path: direct hit (covers non-prefixed attributes and exact-name
    // prefixed ones). getAttribute returns null when absent.
    const direct = this.el.getAttribute(name);
    if (direct !== null) return direct;

    const colonIndex = name.indexOf(":");
    const localName = colonIndex >= 0 ? name.slice(colonIndex + 1) : name;
    const namespaceUri =
      colonIndex >= 0 ? this.resolveAttributeNamespace(name.slice(0, colonIndex)) : undefined;

    for (let i = 0; i < this.el.attributes.length; i++) {
      const attr = this.el.attributes[i];
      if (attr.localName !== localName) continue;
      if (colonIndex < 0) return attr.value;
      if (namespaceUri ? attr.namespaceURI === namespaceUri : attr.namespaceURI !== null) {
        return attr.value;
      }
    }

    return undefined;
  }

  private resolveAttributeNamespace(prefix: string): string | undefined {
    // Optional call: not all DOM implementations (e.g. happy-dom) provide it.
    return (
      this.el?.lookupNamespaceURI?.(prefix) ??
      (prefix === "r"
        ? "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        : undefined)
    );
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
    if (!this.el) return SafeXmlNode.EMPTY;
    const index = getChildIndex(this.el);
    if (index !== undefined) {
      const match = index.get(localName);
      return match === undefined ? SafeXmlNode.EMPTY : new SafeXmlNode(match[0]);
    }
    const children = this.el.children;
    for (let i = 0; i < children.length; i++) {
      if (children[i].localName === localName) {
        return new SafeXmlNode(children[i]);
      }
    }
    return SafeXmlNode.EMPTY;
  }

  /**
   * Get child elements, optionally filtered by localName (namespace-agnostic).
   * If no localName is given, returns all direct child elements.
   */
  children(localName?: string): SafeXmlNode[] {
    if (!this.el) return [];
    if (localName !== undefined) {
      const index = getChildIndex(this.el);
      if (index !== undefined) {
        const match = index.get(localName);
        return match === undefined ? [] : match.map((el) => new SafeXmlNode(el));
      }
    }
    const children = this.el.children;
    const result: SafeXmlNode[] = [];
    for (let i = 0; i < children.length; i++) {
      if (localName === undefined || children[i].localName === localName) {
        result.push(new SafeXmlNode(children[i]));
      }
    }
    return result;
  }

  /** Get the text content, or empty string if the element is missing. */
  text(): string {
    if (!this.el) return "";
    return this.el.textContent ?? "";
  }

  /** Whether the underlying element actually exists. */
  exists(): boolean {
    return this.el !== null;
  }

  /** All direct child elements as SafeXmlNode[]. */
  allChildren(): SafeXmlNode[] {
    return this.children();
  }

  /** The localName of the underlying element, or empty string. */
  get localName(): string {
    return this.el?.localName ?? "";
  }

  /** Raw access to the underlying Element (may be null). */
  get element(): Element | null {
    return this.el;
  }
}

// DOMParser is stateless; construct once instead of per part (a large deck
// parses hundreds of parts).
let sharedParser: DOMParser | undefined;

/**
 * Parse an XML string into a SafeXmlNode wrapping the document element.
 * Uses the browser's built-in DOMParser.
 */
export function parseXml(xmlString: string): SafeXmlNode {
  sharedParser ??= new DOMParser();
  const doc = sharedParser.parseFromString(xmlString, "application/xml");

  // Check for parser errors — DOMParser returns a parsererror document on
  // failure. Browsers differ on placement (document root vs. child of the
  // root), so check the root's tag and a direct tag-name scan; both are far
  // cheaper than a querySelector CSS match on every parsed part.
  const root = doc.documentElement;
  const errorNode =
    root?.localName === "parsererror" ? root : doc.getElementsByTagName("parsererror")[0];
  if (errorNode) {
    console.warn("XML parse error:", errorNode.textContent);
    return SafeXmlNode.EMPTY;
  }

  return new SafeXmlNode(root);
}
