/**
 * The overlay's parts: `SelectionBox`, `ResizeHandle`, `RotateHandle`,
 * `SnapGuide`, and `Marquee` compose inside `Presentation.Selection` to restyle or replace its
 * look, and render the defaults when it has no children.
 */
import * as React from "react";
import { act } from "react";

import { fireEvent, render, waitFor } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";

import { Presentation, useSelection } from "../index";
import type { UseSelectionResult } from "../index";
import { normalizeAngle } from "../selection/geometry";
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

describe("rotating", () => {
  /**
   * Selects shape 2 alone and returns a pointer helper around its center. With
   * layout zeroed in happy-dom, client coordinates are slide pixels times zoom.
   */
  async function renderRotatable() {
    const rendered = await renderOverlay();
    const node = rendered.store.getActiveSlide()!.nodes[0];
    const { zoom } = rendered.store.getState();
    const cx = (node.position.x + node.size.w / 2) * zoom;
    const cy = (node.position.y + node.size.h / 2) * zoom;

    act(() => {
      rendered.overlay.focus();
      fireEvent.keyDown(rendered.overlay, { key: "a", ctrlKey: true });
    });
    const handle = rendered.container.querySelector<HTMLElement>("[data-rotate-handle]")!;

    /** Presses above the center, then drags to `degrees` clockwise from there. */
    function drag(degrees: number, init: { shiftKey?: boolean } = {}) {
      const radians = ((degrees - 90) * Math.PI) / 180;
      act(() => {
        fireEvent.pointerDown(handle, { button: 0, clientX: cx, clientY: cy - 100, pointerId: 1 });
      });
      act(() => {
        fireEvent.pointerMove(rendered.overlay, {
          clientX: cx + 100 * Math.cos(radians),
          clientY: cy + 100 * Math.sin(radians),
          buttons: 1,
          pointerId: 1,
          ...init,
        });
      });
    }

    return { ...rendered, node, handle, drag };
  }

  it("shows a rotate handle on each selected outline", async () => {
    const { container } = await renderRotatable();
    const boxes = container.querySelectorAll("[data-selection-box]");
    expect(boxes).toHaveLength(2);
    for (const box of boxes) {
      expect(box.querySelectorAll("[data-rotate-handle]")).toHaveLength(1);
    }
  });

  it("rotates the selection by the angle dragged around the shape", async () => {
    const { store, overlay, container, drag } = await renderRotatable();
    drag(90);
    expect(overlay.getAttribute("data-mode")).toBe("rotate");

    // The outline follows the drag; resize handles step aside.
    const box = container.querySelector<HTMLElement>("[data-selection-box]")!;
    expect(box.style.transform).toBe("rotate(90deg)");
    expect(box.hasAttribute("data-rotating")).toBe(true);
    expect(container.querySelectorAll("[data-resize-grip]")).toHaveLength(0);

    await act(async () => {
      fireEvent.pointerUp(overlay, { pointerId: 1 });
    });
    const rotations = store.getActiveSlide()!.nodes.map((node) => Math.round(node.rotation));
    expect(rotations).toEqual([90, 90]);

    act(() => {
      store.undo();
    });
    expect(store.getActiveSlide()!.nodes.map((node) => node.rotation)).toEqual([0, 0]);
  });

  it("shows the rotate cursor on the handle and keeps it through the drag", async () => {
    const { overlay, handle, drag } = await renderRotatable();
    expect(handle.style.cursor).toContain("--presentation-rotate-cursor");
    expect(overlay.style.cursor).toBe("");

    drag(45);
    // The drag captures the pointer onto the overlay, which takes the cursor.
    expect(overlay.style.cursor).toContain("--presentation-rotate-cursor");

    await act(async () => {
      fireEvent.pointerUp(overlay, { pointerId: 1 });
    });
    expect(overlay.style.cursor).toBe("");
  });

  it("steps by 15 degrees with Shift held", async () => {
    const { store, overlay, drag } = await renderRotatable();
    drag(50, { shiftKey: true });
    await act(async () => {
      fireEvent.pointerUp(overlay, { pointerId: 1 });
    });
    expect(store.getActiveSlide()!.nodes[0].rotation).toBe(45);
  });

  it("puts the shape back when the drag is cancelled", async () => {
    const { store, overlay, container, drag } = await renderRotatable();
    const shape = container.querySelector<HTMLElement>('[data-pptx-node-id="2"]')!;
    const before = shape.style.transform;

    drag(90);
    expect(shape.style.transform).toContain("rotate(90deg)");
    act(() => {
      fireEvent.pointerCancel(overlay, { pointerId: 1 });
    });

    expect(shape.style.transform).toBe(before);
    expect(overlay.getAttribute("data-mode")).toBe("selected");
    expect(store.getActiveSlide()!.nodes[0].rotation).toBe(0);
  });

  it("keeps the rotate handle on rotated shapes, without resize handles", async () => {
    const { overlay, container, drag } = await renderRotatable();
    drag(30);
    await act(async () => {
      fireEvent.pointerUp(overlay, { pointerId: 1 });
    });

    await waitFor(() => {
      expect(container.querySelectorAll("[data-rotate-handle]")).toHaveLength(2);
    });
    expect(container.querySelectorAll("[data-resize-grip]")).toHaveLength(0);
  });
});

describe("resizing a rotated shape", () => {
  it("shows handles on a lone rotated shape and resizes along its axes", async () => {
    const selection: { current: UseSelectionResult | null } = { current: null };
    function Grab() {
      selection.current = useSelection();
      return null;
    }
    const { store, container, overlay } = await renderOverlay(
      <>
        <Grab />
        <Presentation.SelectionBox>
          <Presentation.ResizeHandle />
        </Presentation.SelectionBox>
      </>,
    );
    const slideId = store.getActiveSlide()!.id;
    await act(async () => {
      await store.edit({ type: "setNodeTransform", slideId, nodeId: "2", rotation: 90 });
    });
    act(() => selection.current!.select(["2"]));

    const handles = container.querySelectorAll<HTMLElement>("[data-resize-grip]");
    expect(handles).toHaveLength(8);
    // The east handle points down the slide now, and its cursor says so.
    const east = container.querySelector<HTMLElement>('[data-resize-grip="e"]')!;
    expect(east.style.cursor).toBe("ns-resize");

    const before = store.getActiveSlide()!.nodes[0];
    const { w, h } = before.size;
    const centerY = before.position.y + h / 2;
    const { zoom } = store.getState();
    act(() => {
      fireEvent.pointerDown(east, { button: 0, clientX: 0, clientY: 0, pointerId: 1 });
    });
    act(() => {
      fireEvent.pointerMove(overlay, { clientX: 0, clientY: 40 * zoom, buttons: 1, pointerId: 1 });
    });
    expect(overlay.style.cursor).toBe("ns-resize");
    await act(async () => {
      fireEvent.pointerUp(overlay, { pointerId: 1 });
    });

    const after = store.getActiveSlide()!.nodes[0];
    expect(after.size.w).toBeCloseTo(w + 40);
    expect(after.size.h).toBeCloseTo(h);
    expect(after.rotation).toBe(90);
    // The west edge, now the top of the turned shape, stays put.
    expect(after.position.y + after.size.h / 2 - after.size.w / 2).toBeCloseTo(centerY - w / 2);
  });
});

describe("normalizeAngle", () => {
  it("wraps angles into [0, 360)", () => {
    expect(normalizeAngle(0)).toBe(0);
    expect(normalizeAngle(360)).toBe(0);
    expect(normalizeAngle(-90)).toBe(270);
    expect(normalizeAngle(450)).toBe(90);
    expect(normalizeAngle(359.9999999)).toBe(0);
  });
});

describe("snap guides", () => {
  it("run across the whole slide, not just along the snapped edge", async () => {
    const selection: { current: UseSelectionResult | null } = { current: null };
    function Grab() {
      selection.current = useSelection();
      return null;
    }

    const { store, container, overlay } = await renderOverlay(
      <>
        <Grab />
        <Presentation.SelectionBox>
          <Presentation.ResizeHandle />
        </Presentation.SelectionBox>
        <Presentation.SnapGuide />
      </>,
    );
    const { zoom, presentation } = store.getState();
    act(() => selection.current!.select(["2"]));

    // Shape 2's right edge sits at 416px and shape 3's left edge at 500px:
    // dragging the east handle 81px lands 3px short, inside the threshold.
    const handle = container.querySelector<HTMLElement>('[data-resize-grip="e"]')!;
    act(() => {
      fireEvent.pointerDown(handle, { button: 0, clientX: 0, clientY: 0, pointerId: 1 });
    });
    act(() => {
      fireEvent.pointerMove(overlay, { clientX: 81 * zoom, clientY: 0, buttons: 1, pointerId: 1 });
    });

    // The resize cursor follows the captured pointer onto the overlay.
    expect(overlay.style.cursor).toBe("ew-resize");

    const guide = container.querySelector<HTMLElement>('[data-snap-guide="x"]')!;
    expect(guide.style.left).toBe(`${500 * zoom}px`);
    expect(guide.style.top).toBe("0px");
    expect(guide.style.height).toBe(`${presentation!.height * zoom}px`);

    await act(async () => {
      fireEvent.pointerUp(overlay, { pointerId: 1 });
    });
    const node = store.getActiveSlide()!.nodes[0];
    expect(node.position.x + node.size.w).toBeCloseTo(500);
  });
});
