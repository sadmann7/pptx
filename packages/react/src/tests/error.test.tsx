import * as React from "react";

import { screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";

import { Presentation } from "../index";
import { createStore } from "../store";
import { loadedStore, withStore } from "./test-utils";

describe("Presentation.Error", () => {
  it("renders the error via render function after a failed load", async () => {
    const store = createStore();
    await store.load(new ArrayBuffer(8)).catch(() => undefined);

    withStore(
      store,
      <Presentation.Error>{(err: Error) => <span>failed: {err.message}</span>}</Presentation.Error>,
    );
    expect(screen.getByText(/^failed: /).textContent).toContain("failed:");
  });

  it("renders nothing when there is no error", async () => {
    const store = await loadedStore();
    const { container } = withStore(store, <Presentation.Error>boom</Presentation.Error>);
    expect(container.textContent).toBe("");
  });
});
