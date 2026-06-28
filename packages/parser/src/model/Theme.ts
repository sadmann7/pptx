import { SafeXmlNode } from "../parser/XmlParser";

export interface ThemeFontInfo {
  latin: string;
  ea: string;
  cs: string;
  scripts?: Record<string, string>;
}

export interface ThemeData {
  colorScheme: Map<string, string>;
  majorFont: ThemeFontInfo;
  minorFont: ThemeFontInfo;
  fillStyles: SafeXmlNode[];
  bgFillStyles?: SafeXmlNode[];
  lineStyles: SafeXmlNode[];
  effectStyles: SafeXmlNode[];
}

const COLOR_SLOTS = [
  "dk1",
  "dk2",
  "lt1",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

function extractColor(node: SafeXmlNode): string | undefined {
  const srgb = node.child("srgbClr");
  if (srgb.exists()) return srgb.attr("val");
  const sys = node.child("sysClr");
  if (sys.exists()) return sys.attr("lastClr") ?? sys.attr("val");
  return undefined;
}

export function parseTheme(root: SafeXmlNode): ThemeData {
  const themeElements = root.child("themeElements");
  const scope = themeElements.exists() ? themeElements : root;

  const clrScheme = scope.child("clrScheme");
  const colorScheme = new Map<string, string>();
  for (const slot of COLOR_SLOTS) {
    const node = clrScheme.child(slot);
    if (node.exists()) {
      const hex = extractColor(node);
      if (hex) colorScheme.set(slot, hex);
    }
  }

  const fontScheme = scope.child("fontScheme");
  const parseFI = (n: SafeXmlNode): ThemeFontInfo => {
    const scripts: Record<string, string> = {};
    for (const f of n.children("font")) {
      const s = f.attr("script"),
        t = f.attr("typeface");
      if (s && t) scripts[s] = t;
    }
    const r: ThemeFontInfo = {
      latin: n.child("latin").attr("typeface") ?? "",
      ea: n.child("ea").attr("typeface") ?? "",
      cs: n.child("cs").attr("typeface") ?? "",
    };
    if (Object.keys(scripts).length > 0) r.scripts = scripts;
    return r;
  };
  const majorFont = parseFI(fontScheme.child("majorFont"));
  const minorFont = parseFI(fontScheme.child("minorFont"));

  const fmtScheme = scope.child("fmtScheme");
  const fillStyles = fmtScheme.child("fillStyleLst").allChildren();
  const bgFillStyles = fmtScheme.child("bgFillStyleLst").allChildren();
  const lineStyles = fmtScheme.child("lnStyleLst").allChildren();
  const effectStyles = fmtScheme.child("effectStyleLst").allChildren();

  return { colorScheme, majorFont, minorFont, fillStyles, bgFillStyles, lineStyles, effectStyles };
}
