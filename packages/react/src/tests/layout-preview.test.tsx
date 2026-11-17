// @vitest-environment jsdom
// jsdom, since happy-dom's XML parser drops namespaced attributes the renderer needs.

import { getLayouts } from "@diceui/pptx-core";
import { describe, expect, it } from "vitest";

import { Presentation } from "../index";
import { editableStore, withStore } from "./test-utils";

describe("LayoutPreview", () => {
  it("renders the layout and releases it on unmount", async () => {
    const store = await editableStore();
    const [layout] = getLayouts(store.getState().presentation!);

    const { container, unmount } = withStore(
      store,
      <Presentation.LayoutPreview layoutId={layout.id} />,
    );

    const preview = container.querySelector(
      "[data-presentation-layout-preview]",
    )!;
    const slideElement = preview.firstElementChild;
    expect(slideElement).toBeTruthy();
    expect(preview.hasAttribute("data-pending")).toBe(false);

    unmount();
    expect(slideElement?.isConnected).toBe(false);
  });

  it("stays pending for a layout the deck does not have", async () => {
    const store = await editableStore();

    const { container } = withStore(
      store,
      <Presentation.LayoutPreview layoutId="ppt/slideLayouts/nope.xml" />,
    );

    const preview = container.querySelector(
      "[data-presentation-layout-preview]",
    )!;
    expect(preview.firstElementChild).toBeNull();
    expect(preview.hasAttribute("data-pending")).toBe(true);
  });
});
