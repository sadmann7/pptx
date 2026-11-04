/**
 * The overlay's parts: `SelectionBox`, `ResizeHandle`, `SnapGuide`, and
 * `Marquee` compose inside `Presentation.Selection` to restyle or replace its
 * look, and render the defaults when it has no children.
 */
import * as React from "react";
import { act } from "react";

import { fireEvent, render, waitFor } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";

import { Presentation } from "../index";
import { editableStore } from "./test-utils";

async function renderOverlay(children?: React.ReactNode) {
  const store = await editableStore();
  const { container } = render(
    <Presentation.Provider store={store}>
      <Presentation.Slide>
        <Presentation.Selection>{children}</Presentation.Selection>
      </Presentation.Slide>
    </Presentation.Provider>,
  );
  await waitFor(() => {
    expect(container.querySelector("[data-pptx-selection]")).not.toBeNull();
  });
  const overlay = container.querySelector<HTMLDivElement>("[data-pptx-selection]")!;

  function selectAll() {
    act(() => {
      overlay.focus();
      fireEvent.keyDown(overlay, { key: "a", ctrlKey: true });
    });
  }

  return { store, container, overlay, selectAll };
}

describe("default parts", () => {
  it("renders a styled outline and eight handles per selected shape", async () => {
    const { container, selectAll } = await renderOverlay();
    selectAll();

    const boxes = container.querySelectorAll<HTMLElement>("[data-selection-box]");
    expect(boxes).toHaveLength(2);
    for (const box of boxes) {
      expect(box.style.boxShadow).toContain("--presentation-selection");
      expect(box.getAttribute("data-mode")).toBe("selected");
      const handles = box.querySelectorAll<HTMLElement>("[data-resize-grip]");
      expect(handles).toHaveLength(8);
      expect(handles[0].style.width).toBe("9px");
    }
  });
});

describe("composed parts", () => {
  it("renders the app's parts unstyled, with class names from state", async () => {
    const { container, selectAll } = await renderOverlay(
      <Presentation.SelectionBox className={({ mode }) => `box box-${mode}`}>
        <Presentation.ResizeHandle className={({ direction }) => `handle handle-${direction}`} />
      </Presentation.SelectionBox>,
    );
    selectAll();

    const boxes = container.querySelectorAll<HTMLElement>("[data-selection-box]");
    expect(boxes).toHaveLength(2);
    expect(boxes[0].className).toBe("box box-selected");
    // Geometry stays inline; the look is the app's.
    expect(boxes[0].style.position).toBe("absolute");
    expect(boxes[0].style.boxShadow).toBe("");

    const handles = boxes[0].querySelectorAll<HTMLElement>("[data-resize-grip]");
    expect(Array.from(handles, (handle) => handle.className)).toEqual([
      "handle handle-nw",
      "handle handle-n",
      "handle handle-ne",
      "handle handle-e",
      "handle handle-se",
      "handle handle-s",
      "handle handle-sw",
      "handle handle-w",
    ]);
    expect(handles[0].style.background).toBe("");
  });

  it("drops the parts an app leaves out", async () => {
    const { container, selectAll } = await renderOverlay(<Presentation.SelectionBox />);
    selectAll();

    expect(container.querySelectorAll("[data-selection-box]")).toHaveLength(2);
    expect(container.querySelectorAll("[data-resize-grip]")).toHaveLength(0);
  });

  it("keeps resizing working from a wrapped handle", async () => {
    const onPointerDown = vi.fn();
    function AppHandle(props: Presentation.ResizeHandle.Props) {
      return <Presentation.ResizeHandle data-testid="app-handle" {...props} />;
    }

    const { store, container, overlay, selectAll } = await renderOverlay(
      <Presentation.SelectionBox>
        <AppHandle onPointerDown={onPointerDown} />
      </Presentation.SelectionBox>,
    );
    const before = store.getActiveSlide()!.nodes[0].size.w;
    selectAll();

    const handle = container.querySelector<HTMLElement>('[data-resize-grip="se"]')!;
    expect(handle.dataset.testid).toBe("app-handle");
    act(() => {
      fireEvent.pointerDown(handle, { button: 0, clientX: 0, clientY: 0, pointerId: 1 });
    });
    expect(onPointerDown).toHaveBeenCalled();
    expect(overlay.getAttribute("data-mode")).toBe("resize");

    act(() => {
      fireEvent.pointerMove(overlay, { clientX: 40, clientY: 40, buttons: 1, pointerId: 1 });
    });
    await act(async () => {
      fireEvent.pointerUp(overlay, { pointerId: 1 });
    });
    expect(store.getActiveSlide()!.nodes[0].size.w).toBeGreaterThan(before);
  });

  it("renders a part through `render` with its state", async () => {
    const { container, selectAll } = await renderOverlay(
      <Presentation.SelectionBox
        render={(props, state) => <section {...props} data-node={state.node.id} />}
      />,
    );
    selectAll();

    const boxes = container.querySelectorAll("section[data-selection-box]");
    expect(Array.from(boxes, (box) => box.getAttribute("data-node"))).toEqual(["2", "3"]);
  });

  it("styles a part with a style function", async () => {
    const { container, selectAll } = await renderOverlay(
      <Presentation.SelectionBox
        style={({ dragging }) => ({ outlineColor: dragging ? "red" : "blue" })}
      />,
    );
    selectAll();

    const box = container.querySelector<HTMLElement>("[data-selection-box]")!;
    expect(box.style.outlineColor).toBe("blue");
    expect(box.style.position).toBe("absolute");
  });
});

describe("part placement", () => {
  it("requires parts to sit inside the overlay", async () => {
    vi.spyOn(console, "error").mockImplementation(() => {});
    const store = await editableStore();

    expect(() =>
      render(
        <Presentation.Provider store={store}>
          <Presentation.SelectionBox />
        </Presentation.Provider>,
      ),
    ).toThrow("`Presentation.SelectionBox` must be used within `Presentation.Selection`.");

    vi.restoreAllMocks();
  });

  it("requires handles to sit inside a selection box", async () => {
    vi.spyOn(console, "error").mockImplementation(() => {});

    await expect(async () => {
      const { selectAll } = await renderOverlay(<Presentation.ResizeHandle />);
      selectAll();
    }).rejects.toThrow(
      "`Presentation.ResizeHandle` must be used within `Presentation.SelectionBox`.",
    );

    vi.restoreAllMocks();
  });
});
