import { describe, expect, it } from "vitest";

import { parseOoxmlBool } from "../parser/booleans";

describe("parseOoxmlBool", () => {
  it("parses xsd:boolean values", () => {
    expect(parseOoxmlBool("1")).toBe(true);
    expect(parseOoxmlBool("true")).toBe(true);
    expect(parseOoxmlBool("0")).toBe(false);
    expect(parseOoxmlBool("false")).toBe(false);
  });

  it("parses ST_OnOff values", () => {
    expect(parseOoxmlBool("on")).toBe(true);
    expect(parseOoxmlBool("t")).toBe(true);
    expect(parseOoxmlBool("off")).toBe(false);
    expect(parseOoxmlBool("f")).toBe(false);
  });

  it("is case- and whitespace-insensitive", () => {
    expect(parseOoxmlBool(" TRUE ")).toBe(true);
    expect(parseOoxmlBool("On")).toBe(true);
  });

  it("returns the default for missing or unrecognized values", () => {
    expect(parseOoxmlBool(undefined)).toBe(false);
    expect(parseOoxmlBool(undefined, true)).toBe(true);
    expect(parseOoxmlBool("maybe")).toBe(false);
    expect(parseOoxmlBool("maybe", true)).toBe(true);
  });
});
