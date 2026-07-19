import * as React from "react";

import { screen, waitFor } from "@testing-library/react";
import { describe, expect, it } from "vitest";

import { Presentation } from "../index";
import { createStore } from "../store";
import { loadedStore, withStore } from "./test-utils";

describe("Presentation.Slide", () => {
  it("mounts the wrapper with data-status and renders slide content when ready", async () => {
    const store = await loadedStore();
    const { container } = withStore(store, <Presentation.Slide data-testid="slide" />);

    const wrapper = screen.getByTestId("slide");
    expect(wrapper.getAttribute("data-status")).toBe("ready");
    await waitFor(() => {
      expect(
        container.querySelector("[data-slide-root], svg, .pptx-slide, div div div"),
      ).not.toBeNull();
    });
  });

  it("keeps the wrapper mounted with data-status=idle before load", () => {
    const store = createStore();
    withStore(store, <Presentation.Slide data-testid="slide" />);
    expect(screen.getByTestId("slide").getAttribute("data-status")).toBe("idle");
  });
});
