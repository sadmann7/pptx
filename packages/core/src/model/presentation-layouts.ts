import type { SafeXmlNode } from "../ooxml/xml";
import type { LayoutData } from "./layout";
import type { PresentationData } from "./presentation";

export interface PresentationLayout {
  /** The layout's part path, a key of `presentation.layouts`. */
  id: string;
  /** The layout's display name, like "Title and Content". */
  name: string;
  /** The part path of the slide master the layout belongs to. */
  masterId: string;
  /** The name of the master's theme, which PowerPoint heads each master's layouts with. */
  themeName?: string;
}

/** Lists a deck's slide layouts grouped by master, in the order PowerPoint's Layout menu shows them. */
export function getLayouts(pres: PresentationData): PresentationLayout[] {
  const ordered = new Set<string>();
  for (const master of pres.masters.values()) {
    for (const layoutId of master.layoutIds ?? []) {
      if (pres.layouts.has(layoutId)) ordered.add(layoutId);
    }
  }
  for (const layoutId of pres.layouts.keys()) ordered.add(layoutId);

  return [...ordered].map((id) => {
    const masterId = pres.layoutToMaster.get(id) ?? "";
    const themeId = pres.masterToTheme.get(masterId);
    return {
      id,
      name:
        pres.layouts.get(id)?.name ??
        id.slice(id.lastIndexOf("/") + 1).replace(/\.xml$/, ""),
      masterId,
      themeName: themeId ? pres.themes.get(themeId)?.name : undefined,
    };
  });
}

export interface PresentationLayoutGroup {
  /** The section divider's layout id, or the master's id for the layouts before any divider. */
  id: string;
  /** The heading for the group: the divider's name, or else the master's theme name. */
  name?: string;
  /** The part path of the slide master the group belongs to. */
  masterId: string;
  /** The group's layouts, in menu order. */
  layouts: PresentationLayout[];
}

/**
 * Groups `getLayouts` by master, and within a master by the section divider layouts some
 * templates use as headings, which are left out of the groups.
 */
export function getLayoutGroups(
  pres: PresentationData,
): PresentationLayoutGroup[] {
  const groups: PresentationLayoutGroup[] = [];
  let current: PresentationLayoutGroup | undefined;
  for (const layout of getLayouts(pres)) {
    const isDivider = isSectionDivider(pres.layouts.get(layout.id));
    if (isDivider || current?.masterId !== layout.masterId) {
      current = {
        id: isDivider ? layout.id : layout.masterId,
        name: isDivider ? layout.name : layout.themeName,
        masterId: layout.masterId,
        layouts: [],
      };
      groups.push(current);
    }
    if (!isDivider) current.layouts.push(layout);
  }
  return groups.filter((group) => group.layouts.length > 0);
}

// A title card such templates insert between sections: no placeholders, and no text but its own name.
function isSectionDivider(layout: LayoutData | undefined): boolean {
  if (!layout?.name || layout.placeholders.length > 0) return false;
  const text = textKey(collectText(layout.spTree));
  return text !== "" && text === textKey(layout.name);
}

function collectText(node: SafeXmlNode): string {
  if (node.localName === "t") return node.text();
  return node
    .allChildren()
    .map((child) => collectText(child))
    .join("");
}

function textKey(text: string): string {
  return text.replace(/\s+/g, "").toLowerCase();
}
