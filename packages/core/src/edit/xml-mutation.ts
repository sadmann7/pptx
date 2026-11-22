import { invalidateChildIndex } from "../ooxml/xml";

export const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const P_NS =
  "http://schemas.openxmlformats.org/presentationml/2006/main";
export const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const RELS_NS =
  "http://schemas.openxmlformats.org/package/2006/relationships";
export const CT_NS =
  "http://schemas.openxmlformats.org/package/2006/content-types";

const DEFAULT_XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

let sharedSerializer: XMLSerializer | undefined;

export function insertChild(
  parent: Element,
  element: Element,
  ref: Element | null,
): void {
  parent.insertBefore(element, ref);
  invalidateChildIndex(parent);
}

export function removeChild(element: Element): Element | null {
  const parent = element.parentElement;
  element.remove();
  if (parent) invalidateChildIndex(parent);
  return parent;
}

export function setOrRemoveAttr(
  element: Element,
  name: string,
  value: string | undefined,
): void {
  if (value === undefined) {
    element.removeAttribute(name);
  } else {
    element.setAttribute(name, value);
  }
}

// XMLSerializer drops the `<?xml ...?>` declaration, so the original one is kept.
export function serializePartText(root: Element, originalText: string): string {
  sharedSerializer ??= new XMLSerializer();
  const declaration =
    /^<\?xml.*?\?>\r?\n?/.exec(originalText)?.[0] ?? DEFAULT_XML_DECLARATION;
  return declaration + sharedSerializer.serializeToString(root);
}
