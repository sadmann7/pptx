import * as React from "react";

import type { SlideHandle } from "@diceui/pptx-core";
import { applySlideScale, renderLayout } from "@diceui/pptx-core";

import { useStoreContext, useStoreSelector } from "./context";
import type { PrimitiveProps } from "./render";
import { renderElement } from "./render";

const LAYOUT_PREVIEW_NAME = "Presentation.LayoutPreview";

export interface LayoutPreviewState {
  /** Part path of the layout being rendered. */
  layoutId: string;
}

export interface LayoutPreviewProps extends PrimitiveProps<
  "div",
  LayoutPreviewState
> {
  /** The layout to render, an `id` from `getLayouts`. */
  layoutId: string;
}

export const LayoutPreview = React.forwardRef<
  HTMLDivElement,
  LayoutPreviewProps
>(function LayoutPreview(
  { layoutId, render, ...layoutPreviewProps },
  forwardedRef,
) {
  const store = useStoreContext(LAYOUT_PREVIEW_NAME);
  const presentation = useStoreSelector(store, (s) => s.presentation, null);
  const layoutPreviewRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const element = layoutPreviewRef.current;
    if (!element || !presentation?.layouts.has(layoutId)) return;

    const mediaUrlCache = new Map<string, string>();
    let slideHandle: SlideHandle | null = null;

    const fit = () => {
      if (slideHandle && element.offsetWidth > 0)
        applySlideScale(
          slideHandle.element,
          element.offsetWidth / presentation.width,
        );
    };

    const show = () => {
      if (slideHandle) return;
      slideHandle = renderLayout(presentation, layoutId, { mediaUrlCache });
      fit();
      element.appendChild(slideHandle.element);
      delete element.dataset.pending;
    };

    const resizeObserver =
      typeof ResizeObserver !== "undefined" ? new ResizeObserver(fit) : null;
    resizeObserver?.observe(element);

    const intersectionObserver =
      typeof IntersectionObserver !== "undefined"
        ? new IntersectionObserver((entries) => {
            if (!entries.some((entry) => entry.isIntersecting)) return;
            intersectionObserver?.disconnect();
            show();
          })
        : null;
    if (intersectionObserver) intersectionObserver.observe(element);
    else show();

    return () => {
      intersectionObserver?.disconnect();
      resizeObserver?.disconnect();
      slideHandle?.element.remove();
      slideHandle?.dispose();
      for (const url of mediaUrlCache.values()) URL.revokeObjectURL(url);
      element.dataset.pending = "";
    };
  }, [presentation, layoutId]);

  return renderElement(
    "div",
    { render },
    {
      state: { layoutId },
      ref: [layoutPreviewRef, forwardedRef],
      props: [
        {
          "aria-hidden": "true",
          "data-presentation-layout-preview": "",
          "data-pending": "",
          inert: true,
          style: {
            width: "100%",
            aspectRatio: `${presentation?.width ?? 16} / ${presentation?.height ?? 9}`,
            overflow: "hidden",
            pointerEvents: "none",
            contain: "layout style paint",
          },
        },
        layoutPreviewProps,
      ],
    },
  );
});

export namespace LayoutPreview {
  export type State = LayoutPreviewState;
  export type Props = LayoutPreviewProps;
}
