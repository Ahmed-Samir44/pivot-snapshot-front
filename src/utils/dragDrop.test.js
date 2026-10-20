import { describe, expect, it } from "vitest";
import { clearDanglingBaseFieldReferences, removeFieldEverywhere } from "./dragDrop";

const base = {
  rows: ["[A]", "[B]"],
  columns: ["[C]", "[Measures]"],
  values: [
    { field: "Revenue", showValuesAs: "PercentOf", baseField: "[C]", baseItem: "x" },
    { field: "Cases", showValuesAs: "RunningTotal", baseField: "[B]", baseItem: "" },
  ],
  filters: [{ field: "[C]", mode: "Members", includedMembers: ["x"] }],
  dateGroupings: [],
  numericGroupings: [],
  calculatedItems: [],
  rowLevelSorts: [{ level: 1, direction: "Ascending" }],
  columnLevelSorts: [{ level: 0, direction: "Ascending" }],
};

describe("clearDanglingBaseFieldReferences", () => {
  it("clears a base field that is no longer on Rows or Columns", () => {
    const values = clearDanglingBaseFieldReferences(base.values, ["[A]", "[B]"]);
    expect(values[0]).toMatchObject({ baseField: null, baseItem: "" });
    expect(values[1].baseField).toBe("[B]");
  });
});

describe("removeFieldEverywhere", () => {
  it("removes a Columns field with its filter, its level sorts and the base fields pointing at it", () => {
    const next = removeFieldEverywhere(base, "[C]");
    expect(next.columns).toEqual(["[Measures]"]);
    expect(next.filters).toEqual([]);
    expect(next.columnLevelSorts).toEqual([]);
    expect(next.rowLevelSorts).toEqual(base.rowLevelSorts);
    expect(next.values[0]).toMatchObject({ baseField: null });
  });
});
