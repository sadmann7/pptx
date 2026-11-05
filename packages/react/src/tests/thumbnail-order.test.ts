import { describe, expect, it } from "vitest";

import { getItemsInPageOrder } from "../thumbnail-list";

describe("getItemsInPageOrder", () => {
  it("orders items as they appear on the page, not as they registered", () => {
    const list = document.createElement("div");
    const [first, added, second] = ["1", "new", "2"].map((id) => {
      const button = document.createElement("button");
      button.dataset.slideId = id;
      list.append(button);
      return button;
    });

    // A slide added after the first one registers last.
    const items = new Map([
      ["1", first],
      ["2", second],
      ["new", added],
    ]);

    expect(getItemsInPageOrder(items)).toEqual([first, added, second]);
  });
});
