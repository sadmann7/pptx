import { describe, expect, it } from "vitest";

import { isAllowedExternalMediaUrl, isAllowedExternalUrl } from "../../utils/url-validation";

describe("isAllowedExternalUrl", () => {
  it("allows http, https, and mailto", () => {
    expect(isAllowedExternalUrl("http://example.com")).toBe(true);
    expect(isAllowedExternalUrl("https://example.com/deck.pptx")).toBe(true);
    expect(isAllowedExternalUrl("mailto:user@example.com")).toBe(true);
  });

  it("rejects script and data URIs", () => {
    expect(isAllowedExternalUrl("javascript:alert(1)")).toBe(false);
    // eslint-disable-next-line oxlint/no-script-url
    expect(isAllowedExternalUrl("JAVASCRIPT:alert(1)")).toBe(false);
    expect(isAllowedExternalUrl("data:text/html,<script>alert(1)</script>")).toBe(false);
    expect(isAllowedExternalUrl("vbscript:msgbox")).toBe(false);
    expect(isAllowedExternalUrl("file:///etc/passwd")).toBe(false);
  });

  it("rejects relative and malformed URLs", () => {
    expect(isAllowedExternalUrl("/relative/path")).toBe(false);
    expect(isAllowedExternalUrl("not a url")).toBe(false);
    expect(isAllowedExternalUrl("")).toBe(false);
  });
});

describe("isAllowedExternalMediaUrl", () => {
  it("allows only http/https, not mailto", () => {
    expect(isAllowedExternalMediaUrl("https://cdn.example.com/img.png")).toBe(true);
    expect(isAllowedExternalMediaUrl("http://cdn.example.com/img.png")).toBe(true);
    expect(isAllowedExternalMediaUrl("mailto:user@example.com")).toBe(false);
    expect(isAllowedExternalMediaUrl("javascript:alert(1)")).toBe(false);
  });
});
