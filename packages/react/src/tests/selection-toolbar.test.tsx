/**
 * An editor toolbar outside the slide: `data-pptx-keep-selection` keeps the
 * selection and text editing alive while it is pressed, and `useSelection`
 * reads and formats what is selected.
 */
import * as React from "react";
import { act } from "react";

import type { ShapeNodeData } from "@diceui/pptx-core";
import { fireEvent, render, waitFor } from "@testing-library/react";
import { describe, expect, it } from "vitest";

import { Presentation, useSelection } from "../index";
import type { UseSelectionResult } from "../index";
import type { Store } from "../store";
import { createStore } from "../store";
import { editableStore } from "./test-utils";

function shapeRuns(store: Store, nodeId: string) {
  const slide = store.getActiveSlide()!;
  const shape = slide.nodes.find((node) => node.id === nodeId) as ShapeNodeData;
  return shape.textBody!.paragraphs[0].runs.map((run) => ({
    text: run.text,
    b: run.properties?.attr("b"),
    i: run.properties?.attr("i"),
  }));
}

async function renderEditor() {
  const store = await editableStore();
  const selection: { current: UseSelectionResult | null } = { current: null };

  function Toolbar() {
    selection.current = useSelection();
    return (
      <div data-pptx-keep-selection="" data-testid="toolbar">
        <button type="button" data-testid="bold">
          Bold
        </button>
      </div>
    );
  }

  const { container, getByTestId } = render(
    <Presentation.Provider store={store}>
      <Toolbar />
      <Presentation.Slide>
        <Presentation.Selection />
      </Presentation.Slide>
    </Presentation.Provider>,
  );

  await waitFor(() => {
    expect(container.querySelector("[data-pptx-selection]")).not.toBeNull();
  });
  const overlay = container.querySelector<HTMLDivElement>("[data-pptx-selection]")!;

  function press(target: Element) {
    act(() => {
      fireEvent.pointerDown(target);
      fireEvent.pointerUp(target);
    });
  }

  /** Selects shape `nodeId` and enters text mode on it, the way Enter does. */
  async function editText(nodeId: string) {
    act(() => selection.current!.select([nodeId]));
    act(() => {
      fireEvent.keyDown(overlay, { key: "Enter" });
    });
    await waitFor(() => expect(overlay.getAttribute("data-mode")).toBe("text"));
    const shape = container.querySelector<HTMLElement>(`[data-pptx-node-id="${nodeId}"]`)!;
    return shape.querySelector<HTMLElement>("[data-pptx-r]")!.firstChild as Text;
  }

  function selectText(node: Node, start: number, end: number) {
    const range = document.createRange();
    range.setStart(node, start);
    range.setEnd(node, end);
    const domSelection = window.getSelection()!;
    domSelection.removeAllRanges();
    domSelection.addRange(range);
    act(() => {
      document.dispatchEvent(new Event("selectionchange"));
    });
  }

  return {
    store,
    container,
    overlay,
    selection,
    toolbar: getByTestId("toolbar"),
    bold: getByTestId("bold"),
    press,
    editText,
    selectText,
  };
}

describe("data-pptx-keep-selection", () => {
  it("keeps the selection when a marked toolbar is pressed", async () => {
    const { overlay, bold, press } = await renderEditor();
    act(() => {
      fireEvent.keyDown(overlay, { key: "a", ctrlKey: true });
    });

    press(bold);
    expect(overlay.getAttribute("data-mode")).toBe("selected");

    press(document.body);
    expect(overlay.getAttribute("data-mode")).toBe("idle");
  });

  it("covers marked content portalled outside the toolbar", async () => {
    const { overlay, press } = await renderEditor();
    act(() => {
      fireEvent.keyDown(overlay, { key: "a", ctrlKey: true });
    });

    const popover = document.createElement("div");
    popover.setAttribute("data-pptx-keep-selection", "");
    const swatch = document.createElement("button");
    popover.appendChild(swatch);
    document.body.appendChild(popover);
    try {
      press(swatch);
      expect(overlay.getAttribute("data-mode")).toBe("selected");
    } finally {
      popover.remove();
    }
  });

  it("keeps text editing alive through toolbar presses and Escape", async () => {
    const { overlay, bold, press, editText } = await renderEditor();
    await editText("2");

    press(bold);
    expect(overlay.getAttribute("data-mode")).toBe("text");

    act(() => {
      fireEvent.keyDown(bold, { key: "Escape" });
    });
    expect(overlay.getAttribute("data-mode")).toBe("text");

    act(() => {
      fireEvent.pointerDown(document.body);
    });
    expect(overlay.getAttribute("data-mode")).toBe("idle");
  });
});

describe("useSelection", () => {
  it("follows the selection and selects or clears it", async () => {
    const { overlay, selection } = await renderEditor();
    expect(selection.current!.mode).toBe("idle");
    expect(selection.current!.nodeIds).toEqual([]);

    act(() => {
      fireEvent.keyDown(overlay, { key: "a", ctrlKey: true });
    });
    expect(selection.current!.nodeIds).toEqual(["2", "3"]);
    expect(selection.current!.nodes.map((node) => node.id)).toEqual(["2", "3"]);

    act(() => selection.current!.select(["3", "missing"]));
    expect(overlay.getAttribute("data-mode")).toBe("selected");
    expect(selection.current!.nodeIds).toEqual(["3"]);

    act(() => selection.current!.clear());
    expect(overlay.getAttribute("data-mode")).toBe("idle");
    expect(selection.current!.mode).toBe("idle");
  });

  it("inserts a shape and selects it", async () => {
    const { store, overlay, selection } = await renderEditor();

    let nodeId: string | undefined;
    await act(async () => {
      nodeId = await selection.current!.insertShape("ellipse");
    });
    const shape = store.getActiveSlide()!.nodes.find((node) => node.id === nodeId) as ShapeNodeData;
    expect(shape.presetGeometry).toBe("ellipse");
    await waitFor(() => expect(selection.current!.nodeIds).toEqual([nodeId]));
    expect(overlay.getAttribute("data-mode")).toBe("selected");
  });

  it("drops a selected copy that undo removed", async () => {
    const { store, selection } = await renderEditor();
    act(() => selection.current!.select(["2"]));

    await act(() => selection.current!.duplicate());
    const [copyId] = selection.current!.nodeIds;
    expect(copyId).not.toBe("2");
    expect(selection.current!.nodes.map((node) => node.id)).toEqual([copyId]);

    await act(() => store.undo());
    expect(selection.current!.nodeIds).toEqual([]);
    expect(selection.current!.nodes).toEqual([]);
  });

  it("formats every selected shape's text as one undo step", async () => {
    const { store, overlay, selection } = await renderEditor();
    act(() => {
      fireEvent.keyDown(overlay, { key: "a", ctrlKey: true });
    });

    await act(() => selection.current!.setTextStyle({ bold: true }));
    expect(shapeRuns(store, "2")).toEqual([{ text: "Slide 1", b: "1", i: undefined }]);
    expect(shapeRuns(store, "3")).toEqual([{ text: "B", b: "1", i: undefined }]);
    await waitFor(() => expect(selection.current!.textStyle?.bold).toBe(true));

    act(() => {
      store.undo();
    });
    expect(shapeRuns(store, "2")[0].b).toBeUndefined();
    expect(shapeRuns(store, "3")[0].b).toBeUndefined();
  });

  it("aligns the selected shapes' paragraphs", async () => {
    const { store, selection } = await renderEditor();
    act(() => selection.current!.select(["2"]));

    await act(() => selection.current!.setParagraphStyle({ align: "center" }));
    const shape = store.getActiveSlide()!.nodes.find((node) => node.id === "2") as ShapeNodeData;
    expect(shape.textBody!.paragraphs[0].properties?.attr("algn")).toBe("ctr");
    await waitFor(() => expect(selection.current!.textStyle?.align).toBe("center"));
  });

  it("formats the selected text in text mode after focus moved to the toolbar", async () => {
    const { store, overlay, selection, bold, press, editText, selectText } = await renderEditor();
    const text = await editText("2");

    selectText(text, 0, 5);
    press(bold);
    // The toolbar took the document selection; the saved range still counts.
    window.getSelection()!.removeAllRanges();

    await act(() => selection.current!.setTextStyle({ italic: true }));
    expect(shapeRuns(store, "2")).toEqual([
      { text: "Slide", b: undefined, i: "1" },
      { text: " 1", b: undefined, i: undefined },
    ]);
    await waitFor(() => expect(overlay.getAttribute("data-mode")).toBe("text"));

    act(() => {
      store.undo();
    });
    expect(shapeRuns(store, "2")).toEqual([{ text: "Slide 1", b: undefined, i: undefined }]);
  });

  it("formats the word at a collapsed caret", async () => {
    const { store, selection, editText, selectText } = await renderEditor();
    const text = await editText("2");

    selectText(text, 2, 2);
    await act(() => selection.current!.setTextStyle({ bold: true }));
    expect(shapeRuns(store, "2")).toEqual([
      { text: "Slide", b: "1", i: undefined },
      { text: " 1", b: undefined, i: undefined },
    ]);
  });

  it("commits pending typing and formats it in the same undo step", async () => {
    const { store, selection, editText, selectText } = await renderEditor();
    const text = await editText("2");

    text.data = "Slide 12";
    selectText(text, 6, 8);
    await act(() => selection.current!.setTextStyle({ bold: true }));
    expect(shapeRuns(store, "2")).toEqual([
      { text: "Slide ", b: undefined, i: undefined },
      { text: "12", b: "1", i: undefined },
    ]);

    act(() => {
      store.undo();
    });
    expect(shapeRuns(store, "2")).toEqual([{ text: "Slide 1", b: undefined, i: undefined }]);
  });

  it("rejects formatting when no Selection is mounted", async () => {
    const store = createStore();
    const selection: { current: UseSelectionResult | null } = { current: null };
    function Toolbar() {
      selection.current = useSelection();
      return null;
    }
    render(
      <Presentation.Provider store={store}>
        <Toolbar />
      </Presentation.Provider>,
    );

    await expect(selection.current!.setTextStyle({ bold: true })).rejects.toThrow(
      /no `Presentation.Selection` is mounted/,
    );
    expect(() => selection.current!.select(["2"])).not.toThrow();
  });
});
