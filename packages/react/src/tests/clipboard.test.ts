import { describe, expect, it } from "vitest";

import {
  readClipboardToken,
  writeClipboardToken,
} from "../selection/clipboard";

function fakeDataTransfer(keep?: Set<string>) {
  const data = new Map<string, string>();
  return {
    setData: (type: string, value: string) => {
      if (!keep || keep.has(type)) data.set(type, value);
    },
    getData: (type: string) => data.get(type) ?? "",
  } as DataTransfer;
}

describe("clipboard token", () => {
  it("round-trips through the custom type", () => {
    const data = fakeDataTransfer();
    writeClipboardToken(data, "abc-123", "One\nTwo <3");

    expect(readClipboardToken(data)).toBe("abc-123");
    expect(data.getData("text/plain")).toBe("One\nTwo <3");
    expect(data.getData("text/html")).toContain("<p>One</p><p>Two &#60;3</p>");
  });

  it("falls back to the HTML when the clipboard drops custom types", () => {
    const data = fakeDataTransfer(new Set(["text/plain", "text/html"]));
    writeClipboardToken(data, "abc-123", "");

    expect(readClipboardToken(data)).toBe("abc-123");
  });

  it("reads nothing from foreign clipboard content", () => {
    const data = fakeDataTransfer();
    data.setData("text/html", "<p>Hello</p>");

    expect(readClipboardToken(data)).toBeNull();
  });
});
