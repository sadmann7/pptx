import React from "react";
import type { TextShape, ThemeColors } from "@pptx/parser";
import { fillToCSS } from "../render/color";
import { bodyStyle } from "../render/text";
import { elementStyle } from "../render/transform";
import { ParagraphElement } from "./shared/ParagraphElement";

interface TextElementProps {
  element: TextShape;
  theme: ThemeColors;
}

export function TextElement({ element, theme }: TextElementProps) {
  const outer: React.CSSProperties = {
    ...elementStyle(element),
    background: fillToCSS(element.fill, theme),
  };

  return (
    <div style={outer} data-element-type="text" data-element-id={element.id}>
      <div style={bodyStyle(element.properties, theme)}>
        {element.paragraphs.map((p, i) => (
          <ParagraphElement key={i} paragraph={p} theme={theme} />
        ))}
      </div>
    </div>
  );
}
