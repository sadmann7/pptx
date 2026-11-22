// The first lookup on an element scans and the second builds the index, so the
// single-pass model build, which queries most elements once, isn't taxed.
// Edits that change children must call invalidateChildIndex.
const childIndexCache = new WeakMap<Element, Map<string, Element[]>>();
const queriedOnce = new WeakSet<Element>();

// Attribute and text mutations don't affect the index.
export function invalidateChildIndex(element: Element): void {
  childIndexCache.delete(element);
  queriedOnce.delete(element);
}

function getChildIndex(element: Element): Map<string, Element[]> | undefined {
  let index = childIndexCache.get(element);
  if (index !== undefined) return index;

  if (!queriedOnce.has(element)) {
    queriedOnce.add(element);
    return undefined;
  }

  index = new Map();
  const children = element.children;
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    const bucket = index.get(child.localName);
    if (bucket === undefined) {
      index.set(child.localName, [child]);
    } else {
      bucket.push(child);
    }
  }
  childIndexCache.set(element, index);
  return index;
}

export class SafeXmlNode {
  readonly element: Element | null;

  constructor(element: Element | null) {
    this.element = element;
  }

  // Every miss shares this instance, since SafeXmlNode is immutable and most
  // optional OOXML elements are absent.
  static readonly EMPTY = new SafeXmlNode(null);

  // An unprefixed name matches any attribute with that local name.
  attr(name: string): string | undefined {
    if (!this.element) return undefined;
    const direct = this.element.getAttribute(name);
    if (direct !== null) return direct;

    const colonIndex = name.indexOf(":");
    const localName = colonIndex >= 0 ? name.slice(colonIndex + 1) : name;
    const namespaceUri =
      colonIndex >= 0
        ? this.resolveAttributeNamespace(name.slice(0, colonIndex))
        : undefined;

    for (let i = 0; i < this.element.attributes.length; i++) {
      const attr = this.element.attributes[i];
      if (attr.localName !== localName) continue;
      if (colonIndex < 0) return attr.value;
      if (
        namespaceUri
          ? attr.namespaceURI === namespaceUri
          : attr.namespaceURI !== null
      ) {
        return attr.value;
      }
    }

    return undefined;
  }

  private resolveAttributeNamespace(prefix: string): string | undefined {
    // Optional call: not all DOM implementations (e.g. happy-dom) provide it.
    return (
      this.element?.lookupNamespaceURI?.(prefix) ??
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
    if (!this.element) return SafeXmlNode.EMPTY;
    const index = getChildIndex(this.element);
    if (index !== undefined) {
      const match = index.get(localName);
      return match === undefined
        ? SafeXmlNode.EMPTY
        : new SafeXmlNode(match[0]);
    }
    const children = this.element.children;
    for (let i = 0; i < children.length; i++) {
      if (children[i].localName === localName) {
        return new SafeXmlNode(children[i]);
      }
    }
    return SafeXmlNode.EMPTY;
  }

  children(localName?: string): SafeXmlNode[] {
    if (!this.element) return [];
    if (localName !== undefined) {
      const index = getChildIndex(this.element);
      if (index !== undefined) {
        const match = index.get(localName);
        return match === undefined
          ? []
          : match.map((element) => new SafeXmlNode(element));
      }
    }
    const children = this.element.children;
    const result: SafeXmlNode[] = [];
    for (let i = 0; i < children.length; i++) {
      if (localName === undefined || children[i].localName === localName) {
        result.push(new SafeXmlNode(children[i]));
      }
    }
    return result;
  }

  text(): string {
    if (!this.element) return "";
    return this.element.textContent ?? "";
  }

  exists(): boolean {
    return this.element !== null;
  }

  allChildren(): SafeXmlNode[] {
    return this.children();
  }

  get localName(): string {
    return this.element?.localName ?? "";
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
    root?.localName === "parsererror"
      ? root
      : doc.getElementsByTagName("parsererror")[0];
  if (errorNode) {
    console.warn("XML parse error:", errorNode.textContent);
    return SafeXmlNode.EMPTY;
  }

  return new SafeXmlNode(root);
}
