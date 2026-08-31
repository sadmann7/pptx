import * as React from "react";

import { render, screen } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";

import { mergeProps, mergeRefs, renderElement } from "../render";

describe("mergeProps", () => {
  it("joins className and shallow-merges style", () => {
    const merged = mergeProps(
      { className: "a", style: { color: "red", margin: 1 } },
      { className: "b", style: { color: "blue" } },
    );
    expect(merged.className).toBe("a b");
    expect(merged.style).toEqual({ color: "blue", margin: 1 });
  });

  it("chains on* handlers, a then b", () => {
    const order: string[] = [];
    const merged = mergeProps(
      { onClick: () => order.push("a") },
      { onClick: () => order.push("b") },
    );
    (merged.onClick as () => void)();
    expect(order).toEqual(["a", "b"]);
  });

  it("lets b overwrite other keys", () => {
    expect(mergeProps({ id: "a", role: "status" }, { id: "b" })).toEqual({
      id: "b",
      role: "status",
    });
  });
});

describe("mergeRefs", () => {
  it("writes through to object and callback refs", () => {
    const objectRef = React.createRef<HTMLDivElement>();
    const callback = vi.fn();
    const node = document.createElement("div");

    mergeRefs(objectRef, callback, null, undefined)(node);

    expect(objectRef.current).toBe(node);
    expect(callback).toHaveBeenCalledWith(node);
  });
});

describe("renderElement", () => {
  it("renders the default tag with composed props", () => {
    render(renderElement("div", {}, { state: {}, props: { role: "status", children: "idle" } }));
    expect(screen.getByRole("status").textContent).toBe("idle");
  });

  it("calls a function render with composed props and state", () => {
    render(
      renderElement(
        "div",
        {
          render: (props, state: { progress: number }) => (
            <span {...props}>at {state.progress}%</span>
          ),
        },
        { state: { progress: 40 }, props: { role: "status" } },
      ),
    );
    expect(screen.getByRole("status").textContent).toBe("at 40%");
  });

  it("clones a ReactElement render and lets its props win", () => {
    render(
      renderElement(
        "div",
        { render: <section data-slot="custom">replaced</section> },
        { state: {}, props: { "data-slot": "default", role: "region" } },
      ),
    );
    const el = screen.getByRole("region");
    expect(el.tagName).toBe("SECTION");
    expect(el.getAttribute("data-slot")).toBe("custom");
    expect(el.textContent).toBe("replaced");
  });

  it("composes className and style from componentProps on top of params", () => {
    render(
      renderElement(
        "div",
        { className: "user", style: { color: "blue" } },
        {
          state: {},
          props: { className: "internal", style: { color: "red", margin: 4 }, children: "x" },
        },
      ),
    );
    const el = screen.getByText("x");
    expect(el.className).toBe("internal user");
    expect(el.style.color).toBe("blue");
    expect(el.style.margin).toBe("4px");
  });

  it("skips undefined entries when merging a props array", () => {
    render(
      renderElement(
        "div",
        {},
        {
          state: {},
          props: [{ role: "note" }, undefined, { children: "kept" }],
        },
      ),
    );
    expect(screen.getByRole("note").textContent).toBe("kept");
  });

  it("defaults button type and img alt, and lets props override them", () => {
    const { rerender, container } = render(
      renderElement("button", {}, { state: {}, props: { children: "Go" } }),
    );
    expect((screen.getByRole("button") as HTMLButtonElement).type).toBe("button");

    rerender(renderElement("button", {}, { state: {}, props: { type: "submit", children: "Go" } }));
    expect((screen.getByRole("button") as HTMLButtonElement).type).toBe("submit");

    rerender(renderElement("img", {}, { state: {}, props: { src: "x.png" } }));
    expect(container.querySelector("img")?.getAttribute("alt")).toBe("");

    rerender(renderElement("img", {}, { state: {}, props: { src: "x.png", alt: "chart" } }));
    expect(container.querySelector("img")?.getAttribute("alt")).toBe("chart");
  });

  it("merges param refs with a render element's ref", () => {
    const paramRef = React.createRef<HTMLButtonElement>();
    const renderRef = React.createRef<HTMLButtonElement>();

    render(
      renderElement(
        "button",
        { render: <button type="button" ref={renderRef} /> },
        { state: {}, ref: paramRef, props: { children: "Go" } },
      ),
    );

    const node = screen.getByRole("button");
    expect(paramRef.current).toBe(node);
    expect(renderRef.current).toBe(node);
  });

  it("warns when a function render looks like a component", () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});

    function LooksLikeComponent(props: React.HTMLAttributes<HTMLDivElement>) {
      return <div {...props} />;
    }

    render(
      renderElement("div", { render: LooksLikeComponent }, { state: {}, props: { children: "x" } }),
    );
    expect(warn).toHaveBeenCalledOnce();
    expect(String(warn.mock.calls[0]?.[0])).toContain("LooksLikeComponent");

    warn.mockClear();
    render(
      renderElement(
        "div",
        {
          render: function renderSpan(props: React.HTMLAttributes<HTMLSpanElement>) {
            return <span {...props} />;
          },
        },
        { state: {}, props: { children: "ok" } },
      ),
    );
    expect(warn).not.toHaveBeenCalled();
    warn.mockRestore();
  });
});
