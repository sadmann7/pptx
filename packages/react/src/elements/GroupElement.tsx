import React from "react";
import type { GroupShape, ThemeColors } from "@pptx/parser";
import { elementStyle } from "../render/transform";
import { SlideElementRenderer } from "./index";

interface GroupElementProps {
  element: GroupShape;
  theme: ThemeColors;
}

export function GroupElement({ element, theme }: GroupElementProps) {
  const outer: React.CSSProperties = {
    ...elementStyle(element),
  };

  return (
    <div style={outer} data-element-type="group" data-element-id={element.id}>
      {element.children.map((child) => (
        <SlideElementRenderer key={child.id} element={child} theme={theme} />
      ))}
    </div>
  );
}
