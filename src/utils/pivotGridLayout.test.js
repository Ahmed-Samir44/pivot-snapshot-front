import { describe, expect, it } from "vitest";
import { applyCollapse, buildDisplayRows, buildRowHeaderGroups, collapseKey, collapsibleKeys, computeVisibleLevels } from "./pivotGridLayout";

const detail = (...labels) => ({ labels, isTotal: false });
const subtotal = (labels, captionLevel, caption) => ({ labels, isTotal: true, captionLevel, caption });
const grand = (labels, caption = "Grand Total") => ({ labels, isTotal: true, captionLevel: -1, caption });
const texts = (row) => row.groupCells.map((c) => c.lines.map((l) => l.text).join("|"));

describe("computeVisibleLevels / buildRowHeaderGroups", () => {
  it("hides an outer label repeated from the row above", () => {
    expect(computeVisibleLevels(["Egypt", "ER"], ["Egypt", "ICU"])).toEqual([false, true]);
    expect(computeVisibleLevels(["KSA", "ER"], ["Egypt", "ER"])).toEqual([true, true]);
    expect(computeVisibleLevels(["Egypt"], null)).toEqual([true]);
  });

  it("merges compact levels into one column", () => {
    expect(buildRowHeaderGroups(3, [true, true, true])).toEqual([{ start: 0, end: 2 }]);
    expect(buildRowHeaderGroups(3, [false, true, false])).toEqual([{ start: 0, end: 0 }, { start: 1, end: 2 }]);
  });
});

describe("buildDisplayRows — tabular, subtotals at the bottom", () => {
  const rows = [
    detail("Egypt", "ICU"),
    detail("Egypt", "ER"),
    subtotal(["Egypt", "Subtotal"], 0, "Egypt Total"),
    detail("KSA", "ICU"),
    subtotal(["KSA", "Subtotal"], 0, "KSA Total"),
    grand(["Total", "Total"]),
  ];
  const groups = buildRowHeaderGroups(2, [false, false]);
  const display = buildDisplayRows(rows, { rowHeaderGroups: groups });

  it("draws details, the item subtotal and one Grand Total row", () => {
    expect(display.map((r) => r.kind)).toEqual(["detail", "detail", "subtotal", "detail", "subtotal", "grand"]);
    expect(texts(display[0])).toEqual(["Egypt", "ICU"]);
    // Egypt isn't repeated on its second row.
    expect(texts(display[1])).toEqual(["", "ER"]);
    expect(texts(display[2])).toEqual(["Egypt Total", ""]);
    expect(texts(display[5])).toEqual(["Grand Total", ""]);
  });

  it("collapses an outer item to its subtotal row", () => {
    const keys = collapsibleKeys(display);
    const egypt = collapseKey(0, ["Egypt"]);
    expect(keys.has(egypt)).toBe(true);
    const collapsed = applyCollapse(display, new Set([egypt]), groups);
    expect(collapsed.map((r) => r.kind)).toEqual(["collapsed", "detail", "subtotal", "grand"]);
    expect(collapsed[0].sourceIndex).toBe(2);
  });
});

describe("Σ Values in Rows", () => {
  // Rows = [Country, Σ]: each country has a row per measure, and a subtotal row per measure.
  const rows = [
    detail("Egypt", "Revenue"),
    detail("Egypt", "Cases"),
    subtotal(["Egypt", "Revenue"], 0, "Egypt Total"),
    subtotal(["Egypt", "Cases"], 0, "Egypt Total"),
    grand(["Total", "Revenue"], "Total Revenue"),
    grand(["Total", "Cases"], "Total Cases"),
  ];
  const groups = buildRowHeaderGroups(2, [false, false]);
  const display = buildDisplayRows(rows, { rowHeaderGroups: groups, measuresLevel: 1 });

  it("doesn't repeat the measure on a Grand Total row whose caption already names it", () => {
    const grands = display.filter((r) => r.kind === "grand");
    expect(grands.map(texts)).toEqual([
      ["Total Revenue", ""],
      ["Total Cases", ""],
    ]);
  });

  it("keeps one row per measure when the item is collapsed", () => {
    const collapsed = applyCollapse(display, new Set([collapseKey(0, ["Egypt"])]), groups);
    const egyptRows = collapsed.filter((r) => r.kind === "collapsed");
    expect(egyptRows.map((r) => r.sourceIndex)).toEqual([2, 3]);
    expect(texts(egyptRows[0])).toEqual(["Egypt", "Revenue"]);
    expect(texts(egyptRows[1])).toEqual(["", "Cases"]);
  });
});
