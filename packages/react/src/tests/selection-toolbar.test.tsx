/**
 * An editor toolbar outside the slide: `data-pptx-keep-selection` keeps the
 * selection and text editing alive while it is pressed, and `useSelection`
 * reads and formats what is selected.
 */
import * as React from "react";
import { act } from "react";

import type { ShapeNodeData, TableNodeData } from "@diceui/pptx-core";
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

function cellTexts(store: Store, nodeId: string) {
  const table = store.getActiveSlide()!.nodes.find((node) => node.id === nodeId) as TableNodeData;
  return table.rows.map((row) =>
    row.cells.map((cell) =>
      cell.textBody!.paragraphs.map((p) => p.runs.map((run) => run.text).join("")).join("\n"),
    ),
  );
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
      <Presentation.Root>
        <Toolbar />
        <Presentation.Slide>
          <Presentation.Selection undoRedoShortcuts />
        </Presentation.Slide>
      </Presentation.Root>
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

describe("anchor", () => {
  it("is the selection's bounds, and null without a selection", async () => {
    const { overlay, selection } = await renderEditor();
    expect(selection.current!.anchor).toBeNull();

    act(() => selection.current!.select(["2"]));
    const anchor = selection.current!.anchor!;
    expect(anchor.contextElement).toBe(overlay);
    expect(anchor.getBoundingClientRect()).toBeInstanceOf(DOMRect);

    act(() => selection.current!.clear());
    expect(selection.current!.anchor).toBeNull();
  });
});

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

    await act(async () => {
      await store.undo();
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

  it("undoes formatting with Ctrl+Z without leaving the text", async () => {
    const { store, overlay, selection, editText, selectText } = await renderEditor();
    const text = await editText("2");
    selectText(text, 0, 5);

    await act(() => selection.current!.setTextStyle({ italic: true }));
    expect(shapeRuns(store, "2")[0]).toEqual({ text: "Slide", b: undefined, i: "1" });

    const editable = overlay.ownerDocument.querySelector<HTMLElement>("[contenteditable]")!;
    await act(async () => {
      fireEvent.keyDown(editable, { key: "z", ctrlKey: true });
    });
    await waitFor(() => expect(shapeRuns(store, "2")[0].i).toBeUndefined());
    expect(overlay.getAttribute("data-mode")).toBe("text");
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

    await act(async () => {
      await store.undo();
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

    await act(async () => {
      await store.undo();
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

describe("tables", () => {
  async function insertTable(editor: Awaited<ReturnType<typeof renderEditor>>) {
    let nodeId: string | undefined;
    await act(async () => {
      nodeId = await editor.selection.current!.insertTable(2, 2);
    });
    await waitFor(() => expect(editor.overlay.getAttribute("data-mode")).toBe("text"));
    return nodeId!;
  }

  function editingCell(container: HTMLElement) {
    const editable = container.querySelector<HTMLElement>(
      "[contenteditable]:not([contenteditable=false])",
    )!;
    return editable.closest("[data-pptx-cell]")!.getAttribute("data-pptx-cell");
  }

  it("inserts a table and edits its first cell", async () => {
    const editor = await renderEditor();
    const nodeId = await insertTable(editor);
    expect(editor.selection.current!.nodeIds).toEqual([nodeId]);
    expect(editingCell(editor.container)).toBe("0,0");
  });

  it("commits a cell's text and moves through cells with Tab and Shift+Tab", async () => {
    const editor = await renderEditor();
    const nodeId = await insertTable(editor);

    const editable = editor.container.querySelector<HTMLElement>("[contenteditable]")!;
    editable.querySelector("[data-pptx-p]")!.textContent = "Name";
    act(() => {
      fireEvent.keyDown(editable, { key: "Tab" });
    });
    await waitFor(() => expect(cellTexts(editor.store, nodeId)[0][0]).toBe("Name"));
    await waitFor(() => expect(editingCell(editor.container)).toBe("0,1"));

    act(() => {
      fireEvent.keyDown(
        editor.container.querySelector<HTMLElement>(
          "[contenteditable]:not([contenteditable=false])",
        )!,
        { key: "Tab", shiftKey: true },
      );
    });
    await waitFor(() => expect(editingCell(editor.container)).toBe("0,0"));
    expect(editor.overlay.getAttribute("data-mode")).toBe("text");
  });

  it("keeps typing that lands before the next cell re-renders", async () => {
    const editor = await renderEditor();
    const nodeId = await insertTable(editor);

    const first = editor.container.querySelector<HTMLElement>("[contenteditable]")!;
    first.querySelector("[data-pptx-p]")!.textContent = "A";
    act(() => {
      fireEvent.keyDown(first, { key: "Tab" });
      const next = editor.container.querySelector<HTMLElement>('[data-pptx-cell="0,1"]')!;
      next.querySelector("[data-pptx-p]")!.textContent = "B";
    });

    await waitFor(() => expect(cellTexts(editor.store, nodeId)[0].slice(0, 2)).toEqual(["A", "B"]));
    await waitFor(() => expect(editingCell(editor.container)).toBe("0,1"));
  });

  it("reports the cell being edited and inserts rows around it", async () => {
    const editor = await renderEditor();
    const nodeId = await insertTable(editor);
    await waitFor(() => expect(editor.selection.current!.cell).toEqual({ row: 0, column: 0 }));

    await act(() => editor.selection.current!.changeTable({ type: "insertRow", side: "above" }));
    await waitFor(() => expect(editingCell(editor.container)).toBe("1,0"));
    expect(cellTexts(editor.store, nodeId)).toHaveLength(3);
    await waitFor(() => expect(editor.selection.current!.cell).toEqual({ row: 1, column: 0 }));

    await act(() => editor.selection.current!.changeTable({ type: "insertColumn", side: "right" }));
    await waitFor(() => expect(cellTexts(editor.store, nodeId)[0]).toHaveLength(3));
    await waitFor(() => expect(editingCell(editor.container)).toBe("1,0"));

    await act(() => editor.selection.current!.changeTable({ type: "deleteRow" }));
    await waitFor(() => expect(cellTexts(editor.store, nodeId)).toHaveLength(2));
    await waitFor(() => expect(editingCell(editor.container)).toBe("1,0"));
  });

  it("adds a row when Tab leaves the last cell", async () => {
    const editor = await renderEditor();
    const nodeId = await insertTable(editor);
    for (const cell of ["0,1", "1,0", "1,1"]) {
      act(() => {
        fireEvent.keyDown(
          editor.container.querySelector<HTMLElement>(
            "[contenteditable]:not([contenteditable=false])",
          )!,
          { key: "Tab" },
        );
      });
      await waitFor(() => expect(editingCell(editor.container)).toBe(cell));
    }

    act(() => {
      fireEvent.keyDown(
        editor.container.querySelector<HTMLElement>(
          "[contenteditable]:not([contenteditable=false])",
        )!,
        { key: "Tab" },
      );
    });
    await waitFor(() => expect(cellTexts(editor.store, nodeId)).toHaveLength(3));
    await waitFor(() => expect(editingCell(editor.container)).toBe("2,0"));
  });

  it("deletes the table with its last row", async () => {
    const editor = await renderEditor();
    let nodeId: string | undefined;
    await act(async () => {
      nodeId = await editor.selection.current!.insertTable(1, 2);
    });
    await waitFor(() => expect(editor.overlay.getAttribute("data-mode")).toBe("text"));

    await act(() => editor.selection.current!.changeTable({ type: "deleteRow" }));
    expect(editor.store.getActiveSlide()!.nodes.some((node) => node.id === nodeId)).toBe(false);
    expect(editor.overlay.getAttribute("data-mode")).toBe("idle");
  });

  it("moves a table from just outside its edge, while selected or editing a cell", async () => {
    const editor = await renderEditor();
    const nodeId = await insertTable(editor);
    const tableOf = () =>
      editor.store.getActiveSlide()!.nodes.find((node) => node.id === nodeId) as TableNodeData;

    for (const mode of ["text", "selected"]) {
      if (mode === "selected") act(() => editor.selection.current!.select([nodeId]));
      await waitFor(() => expect(editor.overlay.getAttribute("data-mode")).toBe(mode));
      const { x } = tableOf().position;
      const frame = editor.container.querySelector<HTMLElement>("[data-border-move]")!;
      act(() => {
        fireEvent.pointerDown(frame, { button: 0, clientX: 100, clientY: 100, pointerId: 1 });
      });
      act(() => {
        fireEvent.pointerMove(editor.overlay, {
          buttons: 1,
          clientX: 140,
          clientY: 100,
          pointerId: 1,
        });
      });
      await act(async () => {
        fireEvent.pointerUp(editor.overlay, { clientX: 140, clientY: 100, pointerId: 1 });
      });
      await waitFor(() => expect(tableOf().position.x).toBeGreaterThan(x));
    }
  });

  it("selects a cell range with Shift+click, formats it, merges it, and splits it", async () => {
    const editor = await renderEditor();
    const nodeId = await insertTable(editor);
    const tableOf = () =>
      editor.store.getActiveSlide()!.nodes.find((node) => node.id === nodeId) as TableNodeData;

    const target = editor.container.querySelector<HTMLElement>('[data-pptx-cell="1,1"]')!;
    act(() => {
      fireEvent.pointerDown(target, { button: 0, shiftKey: true, pointerId: 1 });
    });
    await waitFor(() =>
      expect(editor.selection.current!.cellRange).toEqual({
        start: { row: 0, column: 0 },
        end: { row: 1, column: 1 },
      }),
    );
    expect(editor.container.querySelector("[data-pptx-cell-range]")).not.toBeNull();

    await act(() => editor.selection.current!.setParagraphStyle({ align: "center" }));
    expect(
      tableOf().rows.flatMap((row) =>
        row.cells.map((cell) => cell.textBody!.paragraphs[0].properties?.attr("algn")),
      ),
    ).toEqual(["ctr", "ctr", "ctr", "ctr"]);

    await act(() => editor.selection.current!.changeTable({ type: "mergeCells" }));
    await waitFor(() => expect(tableOf().rows[0].cells[0].gridSpan).toBe(2));
    expect(tableOf().rows[0].cells[0].rowSpan).toBe(2);
    await waitFor(() => expect(editingCell(editor.container)).toBe("0,0"));
    expect(editor.selection.current!.cellRange).toBeNull();

    await act(() => editor.selection.current!.changeTable({ type: "splitCell" }));
    await waitFor(() => expect(tableOf().rows[0].cells[0].gridSpan).toBe(1));
    expect(tableOf().rows[1].cells[1].vMerge).toBe(false);
  });

  it("styles the cell range, the edited cell, or the whole selected table", async () => {
    const editor = await renderEditor();
    const nodeId = await insertTable(editor);
    const fills = () =>
      (
        editor.store.getActiveSlide()!.nodes.find((node) => node.id === nodeId) as TableNodeData
      ).rows.flatMap((row) =>
        row.cells.map((cell) => cell.properties?.child("solidFill").child("srgbClr").attr("val")),
      );

    await act(() => editor.selection.current!.setTableCellStyle({ fill: "#ff0000" }));
    expect(fills()).toEqual(["FF0000", undefined, undefined, undefined]);
    await waitFor(() => expect(editor.overlay.getAttribute("data-mode")).toBe("text"));

    const target = editor.container.querySelector<HTMLElement>('[data-pptx-cell="1,1"]')!;
    act(() => {
      fireEvent.pointerDown(target, { button: 0, shiftKey: true, pointerId: 1 });
    });
    await waitFor(() => expect(editor.selection.current!.cellRange).not.toBeNull());
    await act(() => editor.selection.current!.setTableCellStyle({ fill: "#00ff00" }));
    expect(fills()).toEqual(["00FF00", "00FF00", "00FF00", "00FF00"]);

    act(() => editor.selection.current!.select([nodeId]));
    await act(() => editor.selection.current!.setTableCellStyle({ fill: "#0000ff" }));
    expect(fills()).toEqual(["0000FF", "0000FF", "0000FF", "0000FF"]);
  });

  it("formats every cell of a selected table", async () => {
    const editor = await renderEditor();
    const nodeId = await insertTable(editor);
    act(() => editor.selection.current!.select([nodeId]));

    await act(() => editor.selection.current!.setParagraphStyle({ align: "center" }));
    const table = editor.store
      .getActiveSlide()!
      .nodes.find((node) => node.id === nodeId) as TableNodeData;
    const aligns = table.rows.flatMap((row) =>
      row.cells.map((cell) => cell.textBody!.paragraphs[0].properties?.attr("algn")),
    );
    expect(aligns).toEqual(["ctr", "ctr", "ctr", "ctr"]);
  });
});
