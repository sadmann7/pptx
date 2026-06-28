import React from "react";
import type { ColorMap, TextShape, ThemeColors, ThemeFonts } from "@pptx/parser";
import { effectsToFilter, fillToCSS } from "../render/color";
import { bodyStyle } from "../render/text";
import { elementStyle } from "../render/transform";
import { ParagraphElement } from "./shared/ParagraphElement";

interface TextElementProps {
  element: TextShape;
  theme: ThemeColors;
  colorMap?: ColorMap;
  themeFonts?: ThemeFonts;
}

export function TextElement({ element, theme, colorMap, themeFonts }: TextElementProps) {
  const filter = effectsToFilter(element.effects, theme, colorMap);
  const fontScale = element.properties.fontScale;
  const lnSpcReduction = element.properties.lnSpcReduction;

  const outer: React.CSSProperties = {
    ...elementStyle(element),
    background: fillToCSS(element.fill, theme, colorMap),
    ...(filter ? { filter } : {}),
  };

  return (
    <div style={outer} data-element-type="text" data-element-id={element.id}>
      <div style={bodyStyle(element.properties, theme, colorMap)}>
        {element.paragraphs.map((p, i) => (
          <ParagraphElement
            key={i}
            paragraph={p}
            theme={theme}
            colorMap={colorMap}
            themeFonts={themeFonts}
            fontScale={fontScale}
            lnSpcReduction={lnSpcReduction}
          />
        ))}
      </div>
    </div>
  );
}
