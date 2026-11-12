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
    expect(
      container.querySelector("[data-presentation-selection]"),
    ).not.toBeNull();
  });
  const overlay = container.querySelector<HTMLDivElement>(
    "[data-presentation-selection]",
  )!;

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

    const boxes = container.querySelectorAll<HTMLElement>(
      "[data-presentation-selection-box]",
    );
    expect(boxes).toHaveLength(2);
    for (const box of boxes) {
      expect(box.style.outline).toContain("--presentation-selection");
      expect(box.getAttribute("data-mode")).toBe("selected");
      const handles = box.querySelectorAll<HTMLElement>(
        "[data-presentation-resize-handle]",
      );
      expect(handles).toHaveLength(8);
      expect(handles[0].style.width).toBe("9px");
    }
  });
});

describe("handles on small and flat shapes", () => {
  it("moves handles outside a tiny shape and shows only a flat line's ends", async () => {
    const { store, container, selectAll } = await renderOverlay();
    const slideId = store.getState().activeSlideId!;
    await act(async () => {
      await store.edit({
        type: "batch",
        operations: [
          {
            type: "setNodeTransform",
            slideId,
            nodeId: "2",
            size: { w: 6, h: 6 },
          },
          {
            type: "setNodeTransform",
            slideId,
            nodeId: "3",
            size: { w: 400, h: 0 },
          },
        ],
      });
    });
    selectAll();

    const [tiny, line] = container.querySelectorAll<HTMLElement>(
      "[data-presentation-selection-box]",
    );
    const tinyHandles = tiny.querySelectorAll<HTMLElement>(
      "[data-presentation-resize-handle]",
    );
    expect(tinyHandles).toHaveLength(8);
    const nw = tiny.querySelector<HTMLElement>(
      '[data-presentation-resize-handle="nw"]',
    )!;
    expect(nw.style.transform).toBe(
      "translate(calc(-50% + -12px), calc(-50% + -12px))",
    );

    const lineHandles = Array.from(
      line.querySelectorAll("[data-presentation-resize-handle]"),
    ).map((handle) => handle.getAttribute("data-presentation-resize-handle"));
    expect(lineHandles.sort()).toEqual(["e", "w"]);
  });

  it("keeps handles on the outline of a wide but short shape", async () => {
    const { store, container, selectAll } = await renderOverlay();
    const slideId = store.getState().activeSlideId!;
    await act(async () => {
      await store.edit({
        type: "setNodeTransform",
        slideId,
        nodeId: "2",
        size: { w: 400, h: 10 },
      });
    });
    selectAll();

    const box = container.querySelector<HTMLElement>(
      "[data-presentation-selection-box]",
    )!;
    const nw = box.querySelector<HTMLElement>(
      '[data-presentation-resize-handle="nw"]',
    )!;
    expect(nw.style.transform).toBe(
      "translate(calc(-50% + 0px), calc(-50% + 0px))",
    );
  });
});

describe("composed parts", () => {
  it("renders the app's parts unstyled, with class names from state", async () => {
    const { container, selectAll } = await renderOverlay(
      <Presentation.SelectionBox className={({ mode }) => `box box-${mode}`}>
        <Presentation.ResizeHandle
          className={({ direction }) => `handle handle-${direction}`}
        />
      </Presentation.SelectionBox>,
    );
    selectAll();

    const boxes = container.querySelectorAll<HTMLElement>(
      "[data-presentation-selection-box]",
    );
    expect(boxes).toHaveLength(2);
    expect(boxes[0].className).toBe("box box-selected");
    expect(boxes[0].style.position).toBe("absolute");
    expect(boxes[0].style.boxShadow).toBe("");

    const handles = boxes[0].querySelectorAll<HTMLElement>(
      "[data-presentation-resize-handle]",
    );
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
    const { container, selectAll } = await renderOverlay(
      <Presentation.SelectionBox />,
    );
    selectAll();

    expect(
      container.querySelectorAll("[data-presentation-selection-box]"),
    ).toHaveLength(2);
    expect(
      container.querySelectorAll("[data-presentation-resize-handle]"),
    ).toHaveLength(0);
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

    const handle = container.querySelector<HTMLElement>(
      '[data-presentation-resize-handle="se"]',
    )!;
    expect(handle.dataset.testid).toBe("app-handle");
    act(() => {
      fireEvent.pointerDown(handle, {
        button: 0,
        clientX: 0,
        clientY: 0,
        pointerId: 1,
      });
    });
    expect(onPointerDown).toHaveBeenCalled();
    expect(overlay.getAttribute("data-mode")).toBe("resize");

    act(() => {
      fireEvent.pointerMove(overlay, {
        clientX: 40,
        clientY: 40,
        buttons: 1,
        pointerId: 1,
      });
    });
    await act(async () => {
      fireEvent.pointerUp(overlay, { pointerId: 1 });
    });
    expect(store.getActiveSlide()!.nodes[0].size.w).toBeGreaterThan(before);
  });

  it("renders a part through `render` with its state", async () => {
    const { container, selectAll } = await renderOverlay(
      <Presentation.SelectionBox
        render={(props, state) => (
          <section {...props} data-node={state.node.id} />
        )}
      />,
    );
    selectAll();

    const boxes = container.querySelectorAll(
      "section[data-presentation-selection-box]",
    );
    expect(Array.from(boxes, (box) => box.getAttribute("data-node"))).toEqual([
      "2",
      "3",
    ]);
  });

  it("styles a part with a style function", async () => {
    const { container, selectAll } = await renderOverlay(
      <Presentation.SelectionBox
        style={({ dragging }) => ({ outlineColor: dragging ? "red" : "blue" })}
      />,
    );
    selectAll();

    const box = container.querySelector<HTMLElement>(
      "[data-presentation-selection-box]",
    )!;
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
    ).toThrow(
      "`Presentation.SelectionBox` must be used within `Presentation.Selection`.",
    );

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
  // With layout zeroed in happy-dom, client coordinates are slide pixels times zoom.
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
    const handle = rendered.container.querySelector<HTMLElement>(
      "[data-presentation-rotate-handle]",
    )!;

    function drag(degrees: number, init: { shiftKey?: boolean } = {}) {
      const radians = ((degrees - 90) * Math.PI) / 180;
      act(() => {
        fireEvent.pointerDown(handle, {
          button: 0,
          clientX: cx,
          clientY: cy - 100,
          pointerId: 1,
        });
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
    const boxes = container.querySelectorAll(
      "[data-presentation-selection-box]",
    );
    expect(boxes).toHaveLength(2);
    for (const box of boxes) {
      expect(
        box.querySelectorAll("[data-presentation-rotate-handle]"),
      ).toHaveLength(1);
    }
  });

  it("rotates the selection by the angle dragged around the shape", async () => {
    const { store, overlay, container, drag } = await renderRotatable();
    drag(90);
    expect(overlay.getAttribute("data-mode")).toBe("rotate");

    const box = container.querySelector<HTMLElement>(
      "[data-presentation-selection-box]",
    )!;
    expect(box.style.transform).toBe("rotate(90deg)");
    expect(box.hasAttribute("data-rotating")).toBe(true);
    expect(
      container.querySelectorAll("[data-presentation-resize-handle]"),
    ).toHaveLength(0);

    await act(async () => {
      fireEvent.pointerUp(overlay, { pointerId: 1 });
    });
    const rotations = store
      .getActiveSlide()!
      .nodes.map((node) => Math.round(node.rotation));
    expect(rotations).toEqual([90, 90]);

    await act(async () => {
      await store.undo();
    });
    expect(store.getActiveSlide()!.nodes.map((node) => node.rotation)).toEqual([
      0, 0,
    ]);
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
    const shape = container.querySelector<HTMLElement>(
      '[data-pptx-node-id="2"]',
    )!;
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

  it("keeps rotate and resize handles on rotated shapes in a multi-selection", async () => {
    const { overlay, container, drag } = await renderRotatable();
    drag(30);
    await act(async () => {
      fireEvent.pointerUp(overlay, { pointerId: 1 });
    });

    await waitFor(() => {
      expect(
        container.querySelectorAll("[data-presentation-rotate-handle]"),
      ).toHaveLength(2);
    });
    expect(
      container.querySelectorAll("[data-presentation-resize-handle]"),
    ).toHaveLength(16);
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
      await store.edit({
        type: "setNodeTransform",
        slideId,
        nodeId: "2",
        rotation: 90,
      });
    });
    act(() => selection.current!.select(["2"]));

    const handles = container.querySelectorAll<HTMLElement>(
      "[data-presentation-resize-handle]",
    );
    expect(handles).toHaveLength(8);
    const east = container.querySelector<HTMLElement>(
      '[data-presentation-resize-handle="e"]',
    )!;
    expect(east.style.cursor).toBe("ns-resize");

    const before = store.getActiveSlide()!.nodes[0];
    const { w, h } = before.size;
    const centerY = before.position.y + h / 2;
    const { zoom } = store.getState();
    act(() => {
      fireEvent.pointerDown(east, {
        button: 0,
        clientX: 0,
        clientY: 0,
        pointerId: 1,
      });
    });
    act(() => {
      fireEvent.pointerMove(overlay, {
        clientX: 0,
        clientY: 40 * zoom,
        buttons: 1,
        pointerId: 1,
      });
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
    expect(after.position.y + after.size.h / 2 - after.size.w / 2).toBeCloseTo(
      centerY - w / 2,
    );
  });
});

describe("arrange shortcuts", () => {
  async function renderWithSelection() {
    const selection: { current: UseSelectionResult | null } = { current: null };
    function Grab() {
      selection.current = useSelection();
      return null;
    }
    const rendered = await renderOverlay(
      <>
        <Grab />
        <Presentation.SelectionBox />
      </>,
    );
    const ids = () =>
      rendered.store.getActiveSlide()!.nodes.map((node) => node.id);
    return { ...rendered, selection, ids };
  }

  it("duplicates the selection with Ctrl/Cmd+D and selects the copies", async () => {
    const { store, overlay, selectAll, selection, ids } =
      await renderWithSelection();
    selectAll();

    await act(async () => {
      fireEvent.keyDown(overlay, { key: "d", metaKey: true });
    });
    await waitFor(() => expect(ids()).toEqual(["2", "4", "3", "5"]));
    await waitFor(() => expect(selection.current!.nodeIds).toEqual(["4", "5"]));

    await act(async () => {
      await store.undo();
    });
    expect(ids()).toEqual(["2", "3"]);
  });

  it.each([
    ["BracketRight", true, "front", ["3", "2"]],
    ["BracketRight", false, "forward", ["3", "2"]],
  ] as const)(
    "reorders with Ctrl/Cmd+%s (Shift: %s, %s)",
    async (code, shiftKey, _order, expected) => {
      const { overlay, selection, ids } = await renderWithSelection();
      act(() => selection.current!.select(["2"]));

      await act(async () => {
        fireEvent.keyDown(overlay, { key: "]", code, ctrlKey: true, shiftKey });
      });
      await waitFor(() => expect(ids()).toEqual(expected));
    },
  );

  it("sends a node back with Ctrl/Cmd+Shift+[", async () => {
    const { overlay, selection, ids } = await renderWithSelection();
    act(() => selection.current!.select(["3"]));

    await act(async () => {
      fireEvent.keyDown(overlay, {
        key: "{",
        code: "BracketLeft",
        ctrlKey: true,
        shiftKey: true,
      });
    });
    await waitFor(() => expect(ids()).toEqual(["3", "2"]));
  });
});

describe("clipboard", () => {
  function clipboardEvent(
    type: string,
    data: Map<string, string>,
    files: File[] = [],
  ): Event {
    const event = new Event(type, { bubbles: true, cancelable: true });
    Object.defineProperty(event, "clipboardData", {
      value: {
        setData: (format: string, value: string) => data.set(format, value),
        getData: (format: string) => data.get(format) ?? "",
        files,
      },
    });
    return event;
  }

  async function renderDeck() {
    const selection: { current: UseSelectionResult | null } = { current: null };
    function Grab() {
      selection.current = useSelection();
      return null;
    }
    const rendered = await renderOverlay(
      <>
        <Grab />
        <Presentation.SelectionBox />
      </>,
    );
    const nodes = () => rendered.store.getActiveSlide()!.nodes;
    return { ...rendered, selection, nodes };
  }

  it("copies shapes and pastes them a step along on the same slide", async () => {
    const { overlay, selection, nodes } = await renderDeck();
    act(() => selection.current!.select(["2"]));
    const data = new Map<string, string>();

    act(() => {
      overlay.dispatchEvent(clipboardEvent("copy", data));
    });
    expect(data.get("application/x-diceui-pptx")).toBeTruthy();
    expect(data.get("text/plain")).toBe("Slide 1");

    const original = nodes()[0].position;
    await act(async () => {
      overlay.dispatchEvent(clipboardEvent("paste", data));
    });
    await waitFor(() =>
      expect(nodes().map((node) => node.id)).toEqual(["2", "3", "4"]),
    );
    expect(nodes()[2].position).toEqual({
      x: original.x + 10,
      y: original.y + 10,
    });
    await waitFor(() => expect(selection.current!.nodeIds).toEqual(["4"]));

    await act(async () => {
      overlay.dispatchEvent(clipboardEvent("paste", data));
    });
    await waitFor(() => expect(nodes()).toHaveLength(4));
    expect(nodes()[3].position).toEqual({
      x: original.x + 20,
      y: original.y + 20,
    });
  });

  it("cuts shapes", async () => {
    const { overlay, selection, nodes } = await renderDeck();
    act(() => selection.current!.select(["3"]));
    const data = new Map<string, string>();

    await act(async () => {
      overlay.dispatchEvent(clipboardEvent("cut", data));
    });
    await waitFor(() => expect(nodes().map((node) => node.id)).toEqual(["2"]));

    await act(async () => {
      overlay.dispatchEvent(clipboardEvent("paste", data));
    });
    await waitFor(() => expect(nodes()).toHaveLength(2));
  });

  it("pastes into another deck on the page, where the shape lands in place", async () => {
    const first = await renderDeck();
    const second = await renderDeck();
    act(() => first.selection.current!.select(["2"]));
    const data = new Map<string, string>();

    act(() => {
      first.overlay.dispatchEvent(clipboardEvent("copy", data));
    });
    await act(async () => {
      second.overlay.dispatchEvent(clipboardEvent("paste", data));
    });

    await waitFor(() => expect(second.nodes()).toHaveLength(3));
    expect(first.nodes()).toHaveLength(2);
    expect(second.nodes()[2].position).toEqual(first.nodes()[0].position);
  });

  it("inserts an image file pasted from the system clipboard", async () => {
    const { overlay, selection, nodes } = await renderDeck();
    const png = new Uint8Array(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    new DataView(png.buffer).setUint32(16, 30);
    new DataView(png.buffer).setUint32(20, 20);
    const file = new File([png], "pasted.png", { type: "image/png" });

    await act(async () => {
      overlay.dispatchEvent(clipboardEvent("paste", new Map(), [file]));
    });

    await waitFor(() => expect(nodes()).toHaveLength(3));
    expect(nodes()[2].nodeType).toBe("picture");
    expect(nodes()[2].size).toEqual({ w: 30, h: 20 });
    await waitFor(() =>
      expect(selection.current!.nodeIds).toEqual([nodes()[2].id]),
    );
  });

  it("leaves clipboard events in text fields alone", async () => {
    const { container, selection } = await renderDeck();
    act(() => selection.current!.select(["2"]));
    const input = document.createElement("input");
    container.appendChild(input);
    const data = new Map<string, string>();

    const event = clipboardEvent("copy", data);
    input.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(false);
    expect(data.size).toBe(0);
  });
});

describe("keyboard and screen readers", () => {
  it("moves through the shapes with Tab, then lets focus move on", async () => {
    const { overlay } = await renderOverlay();
    const status = () => overlay.querySelector('[role="status"]')!.textContent;
    const tab = (shiftKey = false) => {
      const event = new KeyboardEvent("keydown", {
        key: "Tab",
        shiftKey,
        bubbles: true,
        cancelable: true,
      });
      act(() => {
        overlay.dispatchEvent(event);
      });
      return event.defaultPrevented;
    };

    act(() => overlay.focus());
    expect(tab()).toBe(true);
    expect(status()).toBe("Rect 1 selected");
    expect(tab()).toBe(true);
    expect(status()).toBe("Rect 1 B selected");

    // Past the last shape, Tab is the browser's again.
    expect(tab()).toBe(false);
    expect(overlay.getAttribute("data-mode")).toBe("idle");
    expect(status()).toBe("");

    expect(tab(true)).toBe(true);
    expect(status()).toBe("Rect 1 B selected");
    expect(tab(true)).toBe(true);
    expect(status()).toBe("Rect 1 selected");
    expect(tab(true)).toBe(false);
  });

  it("announces multi-selections", async () => {
    const { overlay, selectAll } = await renderOverlay();
    selectAll();
    expect(overlay.querySelector('[role="status"]')!.textContent).toBe(
      "2 shapes selected",
    );
    expect(overlay.getAttribute("aria-label")).toBe("Slide editor");
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
    const handle = container.querySelector<HTMLElement>(
      '[data-presentation-resize-handle="e"]',
    )!;
    act(() => {
      fireEvent.pointerDown(handle, {
        button: 0,
        clientX: 0,
        clientY: 0,
        pointerId: 1,
      });
    });
    act(() => {
      fireEvent.pointerMove(overlay, {
        clientX: 81 * zoom,
        clientY: 0,
        buttons: 1,
        pointerId: 1,
      });
    });

    expect(overlay.style.cursor).toBe("ew-resize");

    const guide = container.querySelector<HTMLElement>(
      '[data-presentation-snap-guide="x"]',
    )!;
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
