import { describe, expect, it } from "vitest";
import { createWorkbook, exportPivotToExcel, generalAsShown } from "./exportExcel";
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
    // Egypt's cell carries its own 0.0%; the Grand Total (also 10) keeps the field's General, written
    // as the app shows it (10 → #,##0).
    expect(formats).toEqual(["#,##0.0%", "#,##0"]);
  });

  it("exports the Number options as Excel's own codes (separator, red/parenthesised negatives)", async () => {
    const workbook = createWorkbook();
    await exportPivotToExcel(options({ workbook, valueFields: [{ field: "Revenue", showValuesAs: "Normal", format: { type: "Number", decimalPlaces: 1, useThousandsSeparator: false, negativeNumbers: "RedParentheses" } }] }));
    const formats = new Set();
    workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => cell.value === 10 && formats.add(cell.numFmt)));
    expect([...formats]).toEqual(["0.0;[Red](0.0)"]);
  });

  it("writes General the way the app shows it: separators and only the decimals it needs (up to 2)", () => {
    expect(generalAsShown([2445, 29631])).toBe("#,##0");
    expect(generalAsShown([1564531.1, 7])).toBe("#,##0.0");
    expect(generalAsShown([1.237, NaN, null])).toBe("#,##0.00");
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

  it("with one value field and a Columns field, puts the value name in the corner instead of a repeated row", async () => {
    const workbook = createWorkbook();
    const crossed = {
      rowHeaders: [{ labels: ["AMH"], isTotal: false }],
      columnHeaders: [
        { labels: ["Egypt", "Revenue"], isTotal: false },
        { labels: ["KSA", "Revenue"], isTotal: false },
      ],
      cells: [[1, 2]],
    };
    await exportPivotToExcel(options({ workbook, result: crossed, rows: ["[BU]"], columns: ["[Country]", "[Measures]"], rowFieldLabels: ["BU"], columnFieldLabels: ["Country", "Σ Values"] }));
    const rows = [];
    workbook.worksheets[0].eachRow((row) => rows.push(row.values.slice(1)));
    // (ExcelJS reports a merged range's value on each of its cells.)
    expect(rows[0].slice(0, 2)).toEqual(["Revenue", "Country"]);
    expect(rows.flat().filter((v) => v === "Revenue")).toHaveLength(1);
    expect(rows.flat()).not.toContain("Σ Values");
  });

  it("repeats the row label columns on every printed page, and the header rows with print titles", async () => {
    const workbook = createWorkbook();
    await exportPivotToExcel(options({ workbook, pivotOptions: { ...DEFAULT_PIVOT_OPTIONS, printTitles: true, altTextTitle: "Revenue by country" } }));
    const setup = workbook.worksheets[0].pageSetup;
    expect(setup.printTitlesColumn).toBe("A:A");
    expect(setup.printTitlesRow).toMatch(/^\d+:\d+$/);
    expect(workbook.title).toBe("Revenue by country");
  });

  // 2026-11-10: when the export can't become a real embedded PivotTable (every case above, since
  // targetWorkbook forces that), the sheet is now a real Excel Table (Ctrl+T) — filter buttons and
  // banded rows — instead of bare, unstructured cells (requested live: "عايزه يبقى جدول").
  it("wraps the export in a real Excel Table with a filter button on every column", async () => {
    const workbook = createWorkbook();
    await exportPivotToExcel(options({ workbook }));
    const table = workbook.worksheets[0].getTables()[0];
    expect(table).toBeTruthy();
    expect(table.model.columns.map((c) => c.name)).toEqual(["Country", "Revenue"]);
    expect(table.model.columns.every((c) => c.filterButton)).toBe(true);
    // 3 result rows (Egypt, KSA, Grand Total) = 3 table rows, one each.
    expect(table.model.rows).toHaveLength(3);
  });

  it("flattens a multi-level column header into one combined name per column, dropping the suppressed Σ Values level", async () => {
    const workbook = createWorkbook();
    const crossed = {
      rowHeaders: [{ labels: ["AMH"], isTotal: false }],
      columnHeaders: [
        { labels: ["Egypt", "Revenue"], isTotal: false },
        { labels: ["KSA", "Revenue"], isTotal: false },
      ],
      cells: [[1, 2]],
    };
    await exportPivotToExcel(options({ workbook, result: crossed, rows: ["[BU]"], columns: ["[Country]", "[Measures]"], rowFieldLabels: ["BU"], columnFieldLabels: ["Country", "Σ Values"] }));
    const table = workbook.worksheets[0].getTables()[0];
    // Not "Egypt - Revenue"/"KSA - Revenue": the single-value-field Σ level is suppressed, same as
    // the live grid and the PivotTable-placeholder path (its own "Revenue" caption sits in the
    // corner cell above the table instead, covered by the existing crossed-header test above).
    expect(table.model.columns.map((c) => c.name)).toEqual(["BU", "Egypt", "KSA"]);
  });

  it("keeps a Grand Total column's own caption as its table column name", async () => {
    const withColumnTotal = {
      rowHeaders: [{ labels: ["Egypt"], isTotal: false }],
      columnHeaders: [
        { labels: ["Jan"], isTotal: false },
        { labels: ["Total"], isTotal: true, captionLevel: -1, caption: "Grand Total" },
      ],
      cells: [[5, 5]],
    };
    const workbook = createWorkbook();
    await exportPivotToExcel(options({ workbook, result: withColumnTotal, columns: ["[Month]"], columnFieldLabels: ["Month"] }));
    const table = workbook.worksheets[0].getTables()[0];
    expect(table.model.columns.map((c) => c.name)).toEqual(["Country", "Jan", "Grand Total"]);
  });
});
