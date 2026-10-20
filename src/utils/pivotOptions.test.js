import { describe, expect, it } from "vitest";
import { cellText, DEFAULT_PIVOT_OPTIONS, ERROR_TOKEN, filterAreaSlot, filterValueSummary } from "./pivotOptions";

describe("cellText", () => {
  it("shows Excel's #DIV/0! for an error cell unless custom error text is on", () => {
    expect(cellText(NaN, null)).toBe(ERROR_TOKEN);
    expect(cellText(NaN, null, { ...DEFAULT_PIVOT_OPTIONS, showErrorText: true, errorText: "n/a" })).toBe("n/a");
  });

  it("shows the empty-cell text for a blank cell", () => {
    expect(cellText(null, null)).toBe("");
    expect(cellText(null, null, { ...DEFAULT_PIVOT_OPTIONS, emptyText: "-" })).toBe("-");
    expect(cellText(undefined, null, { ...DEFAULT_PIVOT_OPTIONS, showEmptyText: false, emptyText: "-" })).toBe("");
  });
});

describe("filterAreaSlot", () => {
  it("fills down then over, or over then down, N fields per line", () => {
    const options = { ...DEFAULT_PIVOT_OPTIONS, filterFieldsPerColumn: 2 };
    expect([0, 1, 2].map((i) => filterAreaSlot(i, 3, options))).toEqual([
      { column: 0, row: 0 },
      { column: 0, row: 1 },
      { column: 1, row: 0 },
    ]);
    const across = { ...options, filterAreaOrder: "OverThenDown" };
    expect([0, 1, 2].map((i) => filterAreaSlot(i, 3, across))).toEqual([
      { column: 0, row: 0 },
      { column: 1, row: 0 },
      { column: 0, row: 1 },
    ]);
  });
});

describe("filterValueSummary (the Report Filter caption)", () => {
  it("reads (All), the one item, or (Multiple Items) like Excel", () => {
    expect(filterValueSummary({ mode: "Members", includedMembers: [] })).toBe("(All)");
    expect(filterValueSummary({ mode: "Members", includedMembers: ["Egypt"] })).toBe("Egypt");
    expect(filterValueSummary({ mode: "Members", includedMembers: [""] })).toBe("(blank)");
    expect(filterValueSummary({ mode: "Members", includedMembers: ["A", "B"] })).toBe("(Multiple Items)");
    expect(filterValueSummary({ mode: "LabelContains", includedMembers: [], labelText: "er" })).toBe("(Multiple Items)");
    expect(filterValueSummary({ mode: "DatePeriod", includedMembers: [], labelText: "Today" })).toBe("(Multiple Items)");
  });

  it("never names an EXCLUDED item as the selection", () => {
    // Excluding one item leaves every other item selected.
    expect(filterValueSummary({ mode: "Exclude", includedMembers: [] })).toBe("(All)");
    expect(filterValueSummary({ mode: "Exclude", includedMembers: ["ER"] })).toBe("(Multiple Items)");
  });
});
