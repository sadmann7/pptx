import { parseXml } from "./xml";

export interface RelEntry {
  type: string;
  target: string;
  targetMode?: string;
}

export function isExternalTargetMode(targetMode: string | undefined): boolean {
  return targetMode?.trim().toLowerCase() === "external";
}

function stripUriSuffix(target: string): string {
  const suffixIndex = target.search(/[?#]/);
  return suffixIndex >= 0 ? target.slice(0, suffixIndex) : target;
}

function decodeUriPathSegment(segment: string): string {
  try {
    return decodeURIComponent(segment);
  } catch {
    return segment;
  }
}

/**
 * Parses a .rels XML string into a Map of relationship ID -> RelEntry.
 *
 * Example input:
 * ```xml
 * <Relationships xmlns="...">
 *   <Relationship Id="rId1" Type="http://...slide" Target="slides/slide1.xml"/>
 * </Relationships>
 * ```
 */
export function parseRels(xmlString: string): Map<string, RelEntry> {
  const result = new Map<string, RelEntry>();

  if (!xmlString) return result;

  const root = parseXml(xmlString);
  if (!root.exists()) return result;

  const relationships = root.children("Relationship");
  for (const rel of relationships) {
    const id = rel.attr("Id");
    const type = rel.attr("Type");
    const target = rel.attr("Target");
    const targetMode = rel.attr("TargetMode");

    if (id && type !== undefined && target !== undefined) {
      result.set(id, { type, target, targetMode });
    }
  }

  return result;
}

/**
 * Resolves a relative target path against a base path.
 *
 * ```ts
 * resolveRelTarget("ppt/slides", "../slideLayouts/slideLayout1.xml"); // "ppt/slideLayouts/slideLayout1.xml"
 * resolveRelTarget("ppt/slides", "media/image1.png"); // "ppt/slides/media/image1.png"
 * resolveRelTarget("ppt", "slides/slide1.xml"); // "ppt/slides/slide1.xml"
 * ```
 */
export function resolveRelTarget(basePath: string, target: string): string {
  const targetPath = stripUriSuffix(target);

  // Absolute targets are package-root relative.
  if (targetPath.startsWith("/")) {
    return targetPath.slice(1).replace(/\\/g, "/").split("/").map(decodeUriPathSegment).join("/");
  }

  const baseParts = basePath.replace(/\\/g, "/").split("/").filter(Boolean);
  const targetParts = targetPath
    .replace(/\\/g, "/")
    .split("/")
    .filter(Boolean)
    .map(decodeUriPathSegment);

  const resolved = [...baseParts];
  for (const part of targetParts) {
    if (part === "..") {
      resolved.pop();
    } else if (part !== ".") {
      resolved.push(part);
    }
  }

  return resolved.join("/");
}
