import React from "react";
import type { ColorMap, SlideElement, ThemeColors, ThemeFonts } from "@pptx/parser";
import { TextElement } from "./TextElement";
import { ImageElement } from "./ImageElement";
import { ShapeElement } from "./ShapeElement";
import { TableElement } from "./TableElement";
import { ConnectorElement } from "./ConnectorElement";
import { GroupElement } from "./GroupElement";
import { ChartElement } from "./ChartElement";

export type ElementRendererFn = (
  element: SlideElement,
  theme: ThemeColors,
  colorMap?: ColorMap,
) => React.ReactNode;

interface SlideElementRendererProps {
  element: SlideElement;
  theme: ThemeColors;
  colorMap?: ColorMap;
  themeFonts?: ThemeFonts;
  renderElement?: ElementRendererFn;
}

export function SlideElementRenderer({
  element,
  theme,
  colorMap,
  themeFonts,
  renderElement,
}: SlideElementRendererProps) {
  if (renderElement) {
    const custom = renderElement(element, theme, colorMap);
    if (custom !== undefined) return <>{custom}</>;
  }

  switch (element.type) {
    case "text":
      return (
        <TextElement element={element} theme={theme} colorMap={colorMap} themeFonts={themeFonts} />
      );
    case "image":
      return <ImageElement element={element} theme={theme} colorMap={colorMap} />;
    case "shape":
      return (
        <ShapeElement element={element} theme={theme} colorMap={colorMap} themeFonts={themeFonts} />
      );
    case "table":
      return (
        <TableElement element={element} theme={theme} colorMap={colorMap} themeFonts={themeFonts} />
      );
    case "connector":
      return <ConnectorElement element={element} theme={theme} colorMap={colorMap} />;
    case "group":
      return (
        <GroupElement element={element} theme={theme} colorMap={colorMap} themeFonts={themeFonts} />
      );
    case "chart":
      return <ChartElement element={element} theme={theme} colorMap={colorMap} />;
  }
}

export {
  TextElement,
  ImageElement,
  ShapeElement,
  TableElement,
  ConnectorElement,
  GroupElement,
  ChartElement,
};
