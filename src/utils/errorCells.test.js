import { describe, expect, it } from "vitest";
import { NOT_AVAILABLE, errorToken, isErrorValue, reviveResultCells, serializeResultCells } from "./errorCells";
import { cellText, DEFAULT_PIVOT_OPTIONS } from "./pivotOptions";

// Excel's two cell errors (2026-10-27): #DIV/0! is NaN, #N/A (a missing Show Values As base item) is +Infinity.
describe("error cells", () => {
  it("revives both errors from the API's JSON and writes them back the same way", () => {
    const revived = reviveResultCells({ cells: [["NaN", "Infinity", "-Infinity", 5, null]] });
    expect(Number.isNaN(revived.cells[0][0])).toBe(true);
    expect(revived.cells[0][1]).toBe(NOT_AVAILABLE);
    expect(Number.isNaN(revived.cells[0][2])).toBe(true);
    expect(revived.cells[0].slice(3)).toEqual([5, null]);
    expect(serializeResultCells(revived).cells[0]).toEqual(["NaN", "Infinity", "NaN", 5, null]);
  });

  it("shows each error with Excel's own text, and both as the custom error text", () => {
    expect(isErrorValue(NaN) && isErrorValue(NOT_AVAILABLE)).toBe(true);
    expect(isErrorValue(5)).toBe(false);
    expect(errorToken(NaN)).toBe("#DIV/0!");
    expect(errorToken(NOT_AVAILABLE)).toBe("#N/A");
    expect(cellText(NOT_AVAILABLE, null)).toBe("#N/A");
    expect(cellText(NOT_AVAILABLE, null, { ...DEFAULT_PIVOT_OPTIONS, showErrorText: true, errorText: "-" })).toBe("-");
  });
});
