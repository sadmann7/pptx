import { SafeXmlNode } from "../ooxml/xml";

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
  fillStyles: SafeXmlNode[]; // OOXML style refs index these from 1.
  bgFillStyles?: SafeXmlNode[]; // A bgRef indexes these from 1001.
  lineStyles: SafeXmlNode[]; // OOXML style refs index these from 1.
  effectStyles: SafeXmlNode[]; // OOXML style refs index these from 1.
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
  if (srgb.exists()) {
    return srgb.attr("val");
  }
  const sys = node.child("sysClr");
  if (sys.exists()) {
    return sys.attr("lastClr") ?? sys.attr("val");
  }
  return undefined;
}

// Office themes often leave `<ea>` empty and list script-specific
// `<font script="Hans" />` entries instead, so those are kept too.
function parseFontInfo(fontNode: SafeXmlNode): ThemeFontInfo {
  const scripts: Record<string, string> = {};
  for (const font of fontNode.children("font")) {
    const script = font.attr("script");
    const typeface = font.attr("typeface");
    if (script && typeface) {
      scripts[script] = typeface;
    }
  }

  const result: ThemeFontInfo = {
    latin: fontNode.child("latin").attr("typeface") ?? "",
    ea: fontNode.child("ea").attr("typeface") ?? "",
    cs: fontNode.child("cs").attr("typeface") ?? "",
  };
  if (Object.keys(scripts).length > 0) {
    result.scripts = scripts;
  }
  return result;
}

export function parseTheme(root: SafeXmlNode): ThemeData {
  const themeElements = root.child("themeElements");
  const themeScope = themeElements.exists() ? themeElements : root;

  const clrScheme = themeScope.child("clrScheme");
  const colorScheme = new Map<string, string>();

  for (const slot of COLOR_SLOTS) {
    const slotNode = clrScheme.child(slot);
    if (slotNode.exists()) {
      const hex = extractColor(slotNode);
      if (hex !== undefined) {
        colorScheme.set(slot, hex);
      }
    }
  }

  const fontScheme = themeScope.child("fontScheme");
  const majorFont = parseFontInfo(fontScheme.child("majorFont"));
  const minorFont = parseFontInfo(fontScheme.child("minorFont"));

  const fmtScheme = themeScope.child("fmtScheme");
  const fillStyleLst = fmtScheme.child("fillStyleLst");
  const fillStyles: SafeXmlNode[] = fillStyleLst.allChildren();
  const bgFillStyleLst = fmtScheme.child("bgFillStyleLst");
  const bgFillStyles: SafeXmlNode[] = bgFillStyleLst.allChildren();
  const lnStyleLst = fmtScheme.child("lnStyleLst");
  const lineStyles: SafeXmlNode[] = lnStyleLst.allChildren();
  const effectStyleLst = fmtScheme.child("effectStyleLst");
  const effectStyles: SafeXmlNode[] = effectStyleLst.allChildren();

  return { colorScheme, majorFont, minorFont, fillStyles, bgFillStyles, lineStyles, effectStyles };
}
