import { describe, expect, it } from "vitest";
import { createWorkbook, exportPivotToExcel } from "./exportExcel";
import { DEFAULT_PIVOT_OPTIONS } from "./pivotOptions";
import { cellFormatKey } from "./cellFormats";

// Two countries × one measure, with a Grand Total row — the shape the API returns.
const result = {
  rowHeaders: [
    { labels: ["Egypt"], isTotal: false },
    { labels: ["KSA"], isTotal: false },
    { labels: ["Total"], isTotal: true, captionLevel: -1, caption: "Grand Total" },
  ],
  columnHeaders: [{ labels: ["Revenue"], isTotal: false }],
  cells: [[10], [NaN], [10]],
};

const options = (extra = {}) => ({
  result,
  rows: ["[Country]"],
  columns: ["[Measures]"],
  rowFieldLabels: ["Country"],
  columnFieldLabels: [],
  valueFields: [{ field: "Revenue", showValuesAs: "Normal" }],
  pivotOptions: DEFAULT_PIVOT_OPTIONS,
  ...extra,
});

describe("exportPivotToExcel into a workbook (Show Report Filter Pages)", () => {
  it("adds one sheet per call, with safe and unique names", async () => {
    const workbook = createWorkbook();
    await exportPivotToExcel(options({ workbook, sheetName: "Children's" }));
    await exportPivotToExcel(options({ workbook, sheetName: "Children's" }));
    await exportPivotToExcel(options({ workbook, sheetName: "A/B: [x]" }));

    // An apostrophe would corrupt the print-titles name; duplicates get a suffix; \ / : [ ] are invalid.
    expect(workbook.worksheets.map((s) => s.name)).toEqual(["Children’s", "Children’s (2)", "A B   x"]);
  });

  it("names the sheet after the PivotTable Name unless a sheet name is given", async () => {
    const workbook = createWorkbook();
    await exportPivotToExcel(options({ workbook, pivotOptions: { ...DEFAULT_PIVOT_OPTIONS, name: "Revenue by Doctor" } }));
    await exportPivotToExcel(options({ workbook, pivotOptions: { ...DEFAULT_PIVOT_OPTIONS, name: "  " } }));
    expect(workbook.worksheets.map((s) => s.name)).toEqual(["Revenue by Doctor", "Pivot"]);
  });

  it("gives a cell its own number format (Home › Number) over its value field's", async () => {
    const workbook = createWorkbook();
    const cellFormats = { [cellFormatKey(result.rowHeaders[0], result.columnHeaders[0])]: { type: "Percentage", decimalPlaces: 1 } };
    await exportPivotToExcel(options({ workbook, cellFormats }));
    const formats = [];
    workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => cell.value === 10 && formats.push(cell.numFmt)));
    // Egypt's cell carries its own 0.0%; the Grand Total (also 10) keeps the field's General.
    expect(formats).toEqual(["#,##0.0%", "General"]);
  });

  it("writes real numbers, an error cell as #DIV/0!, and the Grand Total row", async () => {
    const workbook = createWorkbook();
    await exportPivotToExcel(options({ workbook }));
    const values = [];
    workbook.worksheets[0].eachRow((row) => values.push(row.values.slice(1)));
    const flat = values.flat();
    expect(flat).toContain(10);
    expect(flat.some((v) => v?.error === "#DIV/0!" || v === "#DIV/0!")).toBe(true);
    expect(flat).toContain("Grand Total");
  });

  it("repeats the row label columns on every printed page, and the header rows with print titles", async () => {
    const workbook = createWorkbook();
    await exportPivotToExcel(options({ workbook, pivotOptions: { ...DEFAULT_PIVOT_OPTIONS, printTitles: true, altTextTitle: "Revenue by country" } }));
    const setup = workbook.worksheets[0].pageSetup;
    expect(setup.printTitlesColumn).toBe("A:A");
    expect(setup.printTitlesRow).toMatch(/^\d+:\d+$/);
    expect(workbook.title).toBe("Revenue by country");
  });
});
