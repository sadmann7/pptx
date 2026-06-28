import React from "react";
import type { ColorMap, GroupShape, ThemeColors, ThemeFonts } from "@pptx/parser";
import { elementStyle } from "../render/transform";
import { SlideElementRenderer } from "./index";

interface GroupElementProps {
  element: GroupShape;
  theme: ThemeColors;
  colorMap?: ColorMap;
  themeFonts?: ThemeFonts;
}

export function GroupElement({ element, theme, colorMap, themeFonts }: GroupElementProps) {
  const outer: React.CSSProperties = {
    ...elementStyle(element),
    overflow: "hidden",
  };

  const chOff = element.childOffset;
  const chExt = element.childExtent;
  const groupW = element.size.width;
  const groupH = element.size.height;

  const needsRemap =
    chExt && chExt.width > 0 && chExt.height > 0;
  const scaleX = needsRemap ? groupW / chExt.width : 1;
  const scaleY = needsRemap ? groupH / chExt.height : 1;
  const offX = chOff?.x ?? 0;
  const offY = chOff?.y ?? 0;

  // CSS transform remaps child coordinate space (chOff, chExt) to group space (off, ext).
  // Right-to-left: first translate by -chOff (move origin), then scale to group size.
  const innerStyle: React.CSSProperties = needsRemap
    ? {
        position: "absolute",
        left: 0,
        top: 0,
        width: `${chExt.width}pt`,
        height: `${chExt.height}pt`,
        transformOrigin: "0 0",
        transform: `scale(${scaleX}, ${scaleY}) translate(${-offX}pt, ${-offY}pt)`,
      }
    : { position: "absolute", left: 0, top: 0, width: "100%", height: "100%" };

  return (
    <div style={outer} data-element-type="group" data-element-id={element.id}>
      <div style={innerStyle}>
        {element.children.map((child) => (
          <SlideElementRenderer
            key={child.id}
            element={child}
            theme={theme}
            colorMap={colorMap}
            themeFonts={themeFonts}
          />
        ))}
      </div>
    </div>
  );
}
