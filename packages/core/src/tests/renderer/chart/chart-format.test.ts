import { describe, expect, it } from "vitest";

import {
  excelSerialToDateString,
  extractFormatCode,
  extractNumericValues,
  extractNumericValuesWithBlanks,
  extractStringValues,
  formatValue,
} from "../../../renderer/chart/format";
import { parseChartFragment } from "../../fixtures/chart-pptx";

describe("formatValue", () => {
  it("formats General/absent codes with up to 2 decimals", () => {
    expect(formatValue(3, undefined)).toBe("3");
    expect(formatValue(3, "General")).toBe("3");
    expect(formatValue(3.456, undefined)).toBe("3.46");
    expect(formatValue(3.4, "General")).toBe("3.4");
  });

  it("formats percentages by scaling ×100 with the coded decimals", () => {
    expect(formatValue(0.5, "0%")).toBe("50%");
    expect(formatValue(0.256, "0.0%")).toBe("25.6%");
    expect(formatValue(0.12345, "0.00%")).toBe("12.35%");
  });

  it("formats thousands-separated Office number codes", () => {
    expect(formatValue(1234567.891, "#,##0")).toBe("1,234,568");
    expect(formatValue(1234.5, "#,##0.00")).toBe("1,234.50");
  });

  it("uses the negative section with parentheses or minus signs", () => {
    expect(formatValue(-1234.5, "#,##0;(#,##0)")).toBe("(1,235)");
    expect(formatValue(-7, "0;-0")).toBe("-7");
  });

  it("drops trailing zeros for plain decimal codes without grouping", () => {
    // TODO(spec?): Excel renders 2.5 as "2.50" for the "0.00" format code, but
    // the non-grouped path re-parses toFixed output and loses trailing zeros.
    expect(formatValue(2.5, "0.00")).toBe("2.5");
  });

  it("rounds to integers for integer-only codes", () => {
    expect(formatValue(7.6, "0")).toBe("8");
  });
});

describe("excelSerialToDateString", () => {
  it("converts modern serials accounting for the fake 1900 leap day", () => {
    expect(excelSerialToDateString(45292)).toBe("2024/1/1");
  });

  it("converts serials before the 1900 leap-day cutoff without adjustment", () => {
    expect(excelSerialToDateString(59)).toBe("1900/2/28");
  });

  it("returns the raw value for non-date serials", () => {
    expect(excelSerialToDateString(0.5)).toBe("0.5");
    expect(excelSerialToDateString(Number.NaN)).toBe("NaN");
  });
});

describe("extractStringValues", () => {
  it("reads strRef caches preserving pt idx ordering and gaps", () => {
    const cat = parseChartFragment(`<c:cat><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache>
<c:ptCount val="3"/>
<c:pt idx="2"><c:v>C</c:v></c:pt>
<c:pt idx="0"><c:v>A</c:v></c:pt>
</c:strCache></c:strRef></c:cat>`).child("cat");
    expect(extractStringValues(cat)).toEqual(["A", "", "C"]);
  });

  it("falls back to numeric caches, converting date format codes", () => {
    const cat = parseChartFragment(`<c:cat><c:numRef><c:f>Sheet1!$A$1</c:f><c:numCache>
<c:formatCode>m/d/yyyy</c:formatCode>
<c:ptCount val="2"/>
<c:pt idx="0"><c:v>45292</c:v></c:pt>
<c:pt idx="1"><c:v>45323</c:v></c:pt>
</c:numCache></c:numRef></c:cat>`).child("cat");
    expect(extractStringValues(cat)).toEqual(["2024/1/1", "2024/2/1"]);
  });

  it("keeps raw numeric strings when the format code is not date-like", () => {
    const cat = parseChartFragment(`<c:cat><c:numRef><c:numCache>
<c:formatCode>#,##0</c:formatCode>
<c:ptCount val="2"/>
<c:pt idx="0"><c:v>1</c:v></c:pt>
<c:pt idx="1"><c:v>2</c:v></c:pt>
</c:numCache></c:numRef></c:cat>`).child("cat");
    expect(extractStringValues(cat)).toEqual(["1", "2"]);
  });

  it("returns an empty array when no cache exists", () => {
    const cat = parseChartFragment("<c:cat/>").child("cat");
    expect(extractStringValues(cat)).toEqual([]);
  });
});

describe("extractNumericValuesWithBlanks", () => {
  it("fills declared point count, zeroing and flagging missing points", () => {
    const val = parseChartFragment(`<c:val><c:numRef><c:numCache>
<c:ptCount val="3"/>
<c:pt idx="0"><c:v>10</c:v></c:pt>
<c:pt idx="2"><c:v>30</c:v></c:pt>
</c:numCache></c:numRef></c:val>`).child("val");
    const { values, blankIndices } = extractNumericValuesWithBlanks(val);
    expect(values).toEqual([10, 0, 30]);
    expect([...blankIndices]).toEqual([1]);
  });

  it("treats empty and non-numeric point values as blanks", () => {
    const val = parseChartFragment(`<c:val><c:numRef><c:numCache>
<c:ptCount val="2"/>
<c:pt idx="0"><c:v></c:v></c:pt>
<c:pt idx="1"><c:v>oops</c:v></c:pt>
</c:numCache></c:numRef></c:val>`).child("val");
    const { values, blankIndices } = extractNumericValuesWithBlanks(val);
    expect(values).toEqual([0, 0]);
    expect([...blankIndices].sort()).toEqual([0, 1]);
  });

  it("extends past the declared ptCount when higher indices are present", () => {
    const val = parseChartFragment(`<c:val><c:numRef><c:numCache>
<c:ptCount val="1"/>
<c:pt idx="0"><c:v>1</c:v></c:pt>
<c:pt idx="3"><c:v>4</c:v></c:pt>
</c:numCache></c:numRef></c:val>`).child("val");
    expect(extractNumericValues(val)).toEqual([1, 0, 0, 4]);
  });

  it("works on bare numCache containers (no numRef wrapper)", () => {
    const val = parseChartFragment(`<c:val><c:numCache>
<c:ptCount val="1"/>
<c:pt idx="0"><c:v>2.5</c:v></c:pt>
</c:numCache></c:val>`).child("val");
    expect(extractNumericValues(val)).toEqual([2.5]);
  });
});

describe("extractFormatCode", () => {
  it("reads the numCache formatCode", () => {
    const val = parseChartFragment(`<c:val><c:numRef><c:numCache>
<c:formatCode>0.0%</c:formatCode>
<c:ptCount val="0"/>
</c:numCache></c:numRef></c:val>`).child("val");
    expect(extractFormatCode(val)).toBe("0.0%");
  });

  it("returns undefined when the cache or code is missing", () => {
    expect(extractFormatCode(parseChartFragment("<c:val/>").child("val"))).toBeUndefined();
    const noCode = parseChartFragment(
      `<c:val><c:numRef><c:numCache><c:ptCount val="0"/></c:numCache></c:numRef></c:val>`,
    ).child("val");
    expect(extractFormatCode(noCode)).toBeUndefined();
  });
});
