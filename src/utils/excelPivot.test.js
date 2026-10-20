import { describe, expect, it } from "vitest";
import { buildPivotSource } from "./excelPivot";

// Two hospitals × two departments, one measure on Columns, subtotals and a Grand Total — the shape
// the API returns.
const result = {
  rowHeaders: [
    { labels: ["A", "Cardio"] },
    { labels: ["A", "Radio"] },
    { labels: ["A", "Total"], isTotal: true, captionLevel: 0, caption: "A Total" },
    { labels: ["B", "Cardio"] },
    { labels: ["B", "Total"], isTotal: true, captionLevel: 0, caption: "B Total" },
    { labels: ["Total", "Total"], isTotal: true, captionLevel: -1, caption: "Grand Total" },
  ],
  columnHeaders: [{ labels: ["Revenue"] }],
  cells: [[10], [20], [30], [5], [5], [35]],
};
const base = {
  result,
  rows: ["[H]", "[D]"],
  columns: ["[Measures]"],
  rowFieldLabels: ["Hospital", "Department"],
  columnFieldLabels: [],
  valueFields: [{ field: "Revenue", showValuesAs: "Normal" }],
};

describe("buildPivotSource (the exported Excel PivotTable's Data sheet)", () => {
  it("turns every detail cell into a source row, skipping subtotals and the grand total", () => {
    const source = buildPivotSource(base);
    expect(source.ok).toBe(true);
    expect(source.verified).toBe(true);
    expect(source.fields.map((f) => [f.kind, f.name])).toEqual([
      ["row", "Hospital"],
      ["row", "Department"],
      ["data", "Revenue"],
    ]);
    expect(source.data).toEqual([
      ["A", "Cardio", 10],
      ["A", "Radio", 20],
      ["B", "Cardio", 5],
    ]);
    // Items keep the app's order (custom sorts survive), totals as shown in the app.
    expect(source.fields[1].items).toEqual(["Cardio", "Radio"]);
    expect(source.rowGrandTotals).toBe(true);
    expect(source.colGrandTotals).toBe(false);
    expect(source.rowSubtotals).toBe(true);
  });

  it("refuses a measure whose detail cells don't add up to the cube's grand total (distinct count)", () => {
    const cells = [[10], [20], [25], [5], [5], [28]];
    const source = buildPivotSource({ ...base, result: { ...result, cells } });
    expect(source.ok).toBe(false);
    expect(source.reason).toMatch(/Revenue/);
  });

  it("refuses Show Values As and drilldowns", () => {
    expect(buildPivotSource({ ...base, valueFields: [{ field: "Revenue", showValuesAs: "PercentOfGrandTotal" }] }).ok).toBe(false);
    expect(buildPivotSource({ ...base, drillSlot: 0 }).ok).toBe(false);
  });

  it("splits measures and column items into fields, and makes a narrowing report filter a page field", () => {
    const source = buildPivotSource({
      result: {
        rowHeaders: [{ labels: ["A"] }, { labels: ["Total"], isTotal: true, captionLevel: -1 }],
        columnHeaders: [
          { labels: ["2024", "Revenue"] },
          { labels: ["2024", "Cases"] },
          { labels: ["2025", "Revenue"] },
          { labels: ["2025", "Cases"] },
        ],
        cells: [
          [1, 2, 3, 4],
          [1, 2, 3, 4],
        ],
      },
      rows: ["[H]"],
      columns: ["[Year]", "[Measures]"],
      rowFieldLabels: ["Hospital"],
      columnFieldLabels: ["Year", "Σ Values"],
      valueFields: [{ field: "Revenue" }, { field: "Cases" }],
      filters: [
        { field: "[BU]", mode: "Members", includedMembers: ["ASH"] },
        { field: "[Type]", mode: "Members", includedMembers: [] },
      ],
      filterFieldLabels: { "[BU]": "Business Unit" },
    });
    expect(source.ok).toBe(true);
    expect(source.fields.map((f) => f.kind)).toEqual(["page", "row", "col", "data", "data"]);
    expect(source.data).toEqual([
      ["ASH", "A", "2024", 1, 2],
      ["ASH", "A", "2025", 3, 4],
    ]);
    expect(source.measuresInRows).toBe(false);
    expect(source.measuresPosition).toBe(1);
  });
});
