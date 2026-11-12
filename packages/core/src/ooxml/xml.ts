// The first lookup on an element scans and the second builds the index, so the
// single-pass model build, which queries most elements once, isn't taxed.
// Edits that change children must call invalidateChildIndex.
const childIndexCache = new WeakMap<Element, Map<string, Element[]>>();
const queriedOnce = new WeakSet<Element>();

// Attribute and text mutations don't affect the index.
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

  // Every miss shares this instance, since SafeXmlNode is immutable and most
  // optional OOXML elements are absent.
  static readonly EMPTY = new SafeXmlNode(null);

  // An unprefixed name matches any attribute with that local name.
  attr(name: string): string | undefined {
    if (!this.el) return undefined;
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

  numAttr(name: string): number | undefined {
    const raw = this.attr(name);
    if (raw === undefined) return undefined;
    const n = Number(raw);
    return Number.isNaN(n) ? undefined : n;
  }

  // Matches by localName in any namespace and returns an empty node on a miss,
  // so chains never throw.
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

  text(): string {
    if (!this.el) return "";
    return this.el.textContent ?? "";
  }

  exists(): boolean {
    return this.el !== null;
  }

  allChildren(): SafeXmlNode[] {
    return this.children();
  }

  get localName(): string {
    return this.el?.localName ?? "";
  }

  get element(): Element | null {
    return this.el;
  }
}

// DOMParser is stateless; construct once instead of per part (a large deck
// parses hundreds of parts).
let sharedParser: DOMParser | undefined;

export function parseXml(xmlString: string): SafeXmlNode {
  sharedParser ??= new DOMParser();
  // Some generators (e.g. Open XML SDK) write a UTF-8 BOM into XML parts.
  // JSZip's string decoder keeps it as a leading U+FEFF, which strict
  // DOMParser implementations reject, so strip it before parsing.
  if (xmlString.charCodeAt(0) === 0xfeff) {
    xmlString = xmlString.slice(1);
  }
  const doc = sharedParser.parseFromString(xmlString, "application/xml");

  // DOMParser returns a parsererror document on failure, which browsers place
  // at the root or as its child; checking both is cheaper than a querySelector.
  const root = doc.documentElement;
  const errorNode =
    root?.localName === "parsererror" ? root : doc.getElementsByTagName("parsererror")[0];
  if (errorNode) {
    console.warn("XML parse error:", errorNode.textContent);
    return SafeXmlNode.EMPTY;
  }

  return new SafeXmlNode(root);
}
