import { describe, expect, it } from "vitest";
import { selectionStats } from "./selectionStats";

describe("selectionStats (the status bar's Sum/Average/Count)", () => {
  it("is null with nothing selected", () => {
    expect(selectionStats([])).toBeNull();
    expect(selectionStats(null)).toBeNull();
  });

  it("sums and averages the numbers, counting every non-empty cell", () => {
    expect(selectionStats([10, 20, null, 30])).toEqual({ cellCount: 4, count: 3, numericCount: 3, sum: 60, average: 20 });
  });

  it("counts an error cell (#DIV/0!) but leaves it out of Sum and Average", () => {
    const stats = selectionStats([NaN, 4]);
    expect(stats.count).toBe(2);
    expect(stats.sum).toBe(4);
    expect(stats.average).toBe(4);
  });

  it("has no Sum/Average when the range holds no numbers", () => {
    expect(selectionStats([null, NaN])).toEqual({ cellCount: 2, count: 1, numericCount: 0, sum: null, average: null });
  });
});
