import ExcelJS from "exceljs";
import { formatForColumn } from "./numberFormat";
import { buildRowHeaderGroups, buildDisplayRows, buildColumnHeaderRows, computeLabelSpans } from "./pivotGridLayout";
import { DEFAULT_PIVOT_OPTIONS, ERROR_TOKEN, filterAreaSlot, filterValueSummary } from "./pivotOptions";
import { isErrorValue } from "./errorCells";
import { measuresIndexOf, buildCaptionToIndex, measureIndexOf } from "./measuresPlacement";
import { DEFAULT_MIN_COLOR, DEFAULT_MID_COLOR, DEFAULT_MAX_COLOR, DEFAULT_BAR_COLOR } from "./conditionalFormat";
import { buildPivotSource, addPivotTableToXlsx } from "./excelPivot";
import { cellFormatKey, hasCellFormats } from "./cellFormats";

// ExcelJS wants ARGB ("FFRRGGBB"), this app's own ConditionalFormat colors are plain "#RRGGBB" —
// alpha is always opaque (FF), matching every hardcoded color this export already used before
// colors became customizable.
function toArgb(hex) {
  return `FF${hex.replace("#", "").toUpperCase()}`;
}

const SUBTOTAL_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF6E9D8" } }; // ~bg-gold/10
const HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } }; // slate-100

// Same clamp/decimal handling as numberFormat.js's formatNumber, but producing an EXCEL number
// format CODE instead of a display string — the cell keeps its real numeric value (sortable,
// summable in Excel) with Excel doing the formatting, instead of exporting pre-formatted text.
//
// Percentage uses Excel's own real percent operator ('%', not a quoted literal) — our stored
// values are a plain ratio (0.3843, not 38.43, fixed 2026-10-05, see numberFormat.js's own WHY
// comment), exactly what Excel's built-in *100-on-render behavior for '%' expects, so no
// workaround is needed here anymore.
function excelNumberFormat(format) {
  const type = format?.type ?? "General";
  const decimals = Math.min(Math.max(format?.decimalPlaces ?? 2, 0), 10);
  const decimalPart = decimals > 0 ? "." + "0".repeat(decimals) : "";
  // No currency is assumed by default — see numberFormat.js's own WHY comment (an earlier "EGP"
  // default was hardcoded here without ever having been asked for, flagged live 2026-10-06). The
  // quoted-literal prefix (and its separating space) is dropped entirely when empty, not left as a
  // dangling `"" ` that would still show a stray leading space in the exported cell.
  const symbol = (format?.currencySymbol ?? "").replace(/"/g, "");
  const currencyPrefix = symbol ? `"${symbol}" ` : "";
  // Excel's own Negative numbers codes: a second section for negatives ([Red] and/or parentheses).
  const withNegatives = (positive) => {
    const style = format?.negativeNumbers ?? "Minus";
    if (style === "Red") return `${positive};[Red]${positive}`;
    if (style === "Parentheses") return `${positive};(${positive})`;
    if (style === "RedParentheses") return `${positive};[Red](${positive})`;
    return positive;
  };
  if (type === "Number") return withNegatives(`${format?.useThousandsSeparator === false ? "0" : "#,##0"}${decimalPart}`);
  if (type === "Currency") return withNegatives(`${currencyPrefix}#,##0${decimalPart}`);
  if (type === "Percentage") return `#,##0${decimalPart}%`;
  // Excel's own accounting format section syntax (positive;negative;zero): the negative section
  // in parentheses is what gives Accounting its own distinct look from Currency. The explicit
  // third (zero) section matches real Excel's own built-in Accounting format showing a plain dash
  // for zero — without it, a 2-section code falls back to the POSITIVE section for zero ("0.00"
  // instead of "-"), the same mismatch just fixed in NumberFormatter.cs/numberFormat.js (added
  // 2026-10-06, full codebase review).
  if (type === "Accounting") return `${currencyPrefix}#,##0${decimalPart};(${currencyPrefix}#,##0${decimalPart});${currencyPrefix}"-"`;
  // Date/Time/Fraction/Scientific: real Excel format codes applied to the SAME raw number our own
  // NumberFormatter.cs/numberFormat.js already treat as an Excel-style serial — Excel's own date
  // system uses the identical 1899-12-30 epoch, so letting Excel itself render these (rather than
  // exporting pre-formatted text) shows the exact same date/time/fraction/scientific value as the
  // live grid, while keeping the cell a real, sortable number in the exported file.
  if (type === "Date") return "yyyy-mm-dd";
  if (type === "Time") return "hh:mm:ss";
  if (type === "Fraction") return "# ??/??";
  if (type === "Scientific") return `0${decimalPart}E+00`;
  // Text: Excel's own "@" format code only changes how NEW typed input is interpreted, not how an
  // already-numeric cell displays — General is the closer real-world match for an exported numeric
  // cell than "@" would be.
  if (type === "Text") return "General";
  // Special: each preset mask (see numberFormat.js's SPECIAL_MASKS) IS already valid Excel number
  // format syntax as-is ('0' placeholders + literal punctuation) — no translation needed.
  if (type === "Special") {
    const masks = {
      ZipCode: "00000",
      ZipCodePlus4: "00000-0000",
      PhoneNumber: "(000) 000-0000",
      SocialSecurityNumber: "000-00-0000",
    };
    return masks[format?.specialType] ?? masks.ZipCode;
  }
  // Custom: the user typed a real Excel format code directly — pass it straight through so Excel
  // renders it exactly as it would any other custom-formatted cell, instead of re-deriving it
  // through our own best-effort FormatAsCustom parser (which is for THIS app's own grid/snapshot
  // display, not for an exported cell Excel will format itself).
  if (type === "Custom") return format?.customFormatCode?.trim() || "General";
  return "General";
}

// Builds a run-length-encoded merge plan for one header LEVEL: adjacent cells sharing the same
// label collapse into a single merged cell spanning them, same idea as how Excel's own PivotTable
// column headers group repeated parent labels — PivotGrid.jsx doesn't do this on screen (an HTML
// table column-merge there would fight the live search/collapse interactions), but a static
// export has no such constraint, and it reads far better in Excel than 8 columns each repeating
// the same label.
function mergeRuns(labels) {
  const runs = [];
  let i = 0;
  while (i < labels.length) {
    let j = i + 1;
    while (j < labels.length && labels[j] === labels[i]) j += 1;
    runs.push({ start: i, end: j, label: labels[i] });
    i = j;
  }
  return runs;
}

// Splits [0, count) into contiguous runs that exclude every index in `excluded` — used to keep a
// value field's conditional formatting scoped to DETAIL cells only, same "totals never get a
// background" rule as the live grid (see ConditionalFormatting.cs) and detailRangeForColumn/Row's
// own totals exclusion, just expressed as which ranges to actually paint rather than which values
// to average.
function detailRuns(count, excluded) {
  const runs = [];
  let start = null;
  for (let i = 0; i < count; i++) {
    if (excluded.has(i)) {
      if (start !== null) runs.push({ start, end: i });
      start = null;
    } else if (start === null) {
      start = i;
    }
  }
  if (start !== null) runs.push({ start, end: count });
  return runs;
}

function excelConditionalFormatRule(format) {
  if (!format || format.type === "None") return null;
  if (format.type === "ColorScale") {
    return {
      type: "colorScale",
      cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }],
      color: [
        { argb: toArgb(format.minColor || DEFAULT_MIN_COLOR) },
        { argb: toArgb(format.midColor || DEFAULT_MID_COLOR) },
        { argb: toArgb(format.maxColor || DEFAULT_MAX_COLOR) },
      ],
    };
  }
  if (format.type === "DataBar") {
    return {
      type: "dataBar",
      cfvo: [{ type: "min" }, { type: "max" }],
      color: { argb: toArgb(format.barColor || DEFAULT_BAR_COLOR) },
      minLength: 0,
      maxLength: 100,
    };
  }
  if (format.type === "IconSet") {
    return {
      type: "iconSet",
      iconSet: "3TrafficLights1",
      cfvo: [{ type: "percent", value: 0 }, { type: "percent", value: 33 }, { type: "percent", value: 67 }],
      showValue: true,
    };
  }
  return null;
}

// Builds and downloads a REAL .xlsx (not an HTML file wearing an .xls extension) from the same
// PivotResult shape PivotGrid.jsx renders — real numeric cells with Excel number-format codes,
// merged headers, and native Excel conditional formatting (color scale/data bar/icon set) rather
// than a flat color painted once at export time, so it stays "live" if the numbers are edited in
// Excel afterward. See DECISIONS.md (2026-09-23) for why this replaced the simpler "export the
// saved HTML as .xls" idea — that approach triggers Excel's "format doesn't match extension"
// warning and can't produce real numeric cells or native conditional formatting at all.
export async function exportPivotToExcel({
  result,
  rows = [],
  columns = [],
  rowFieldLabels = [],
  columnFieldLabels = [],
  valueFields = [],
  filename = "Pivot",
  // Same layout inputs as PivotGrid (2026-10-14) so the export draws the SAME rows: compact-form
  // item rows, "ER Total"/"Grand Total" captions, per-field Compact/Repeat/Blank Line. Omitted
  // (an older caller) = Tabular with no per-field overrides.
  layout = "Tabular",
  fieldLayouts = [],
  repeatItemLabels = false,
  // A per-member drilldown pivot (PLAN 8): ragged rows in ONE indented "Row Labels" column, same
  // as the live grid and the saved snapshot.
  drilldown = false,
  // The report Subtotals setting applied to the drilldown hierarchy (see buildDisplayRows).
  drilldownSubtotals = "Top",
  // A drilldown COMBINED with other Rows fields (Drilldown v3): its slot's index in `rows` (-1 =
  // none). Laid out like any Rows field; buildDisplayRows indents the slot by member depth.
  drillSlot = -1,
  // Columns drilldown (C4): its slot in `columns` (-1 = none) — header rows and spans come from
  // buildColumnHeaderRows, the same as the live grid and the saved snapshot.
  columnsDrillSlot = -1,
  // PivotTable Options > Layout & Format (2026-10-18): empty/error cell text, merged labels, and
  // the report filter area's layout — same as the live grid and the saved snapshot.
  pivotOptions = DEFAULT_PIVOT_OPTIONS,
  // The report filters, written above the table like Excel's Report Filter area, captioned by
  // filterFieldLabels[field].
  filters = [],
  filterFieldLabels = {},
  // Show Report Filter Pages (2026-10-20): several pivots as sheets of ONE workbook — pass the
  // workbook (and a sheet name) and it's only added to, not downloaded; see downloadWorkbook.
  workbook: targetWorkbook = null,
  // The sheet's name — by default the PivotTable Name (2026-10-20, ribbon/Options), so the name the
  // user gave the pivot is what they see on the sheet tab. Report Filter Pages pass the item instead.
  sheetName = null,
  // A real Excel PivotTable over a Data sheet (2026-10-20, see excelPivot.js) — false keeps plain
  // cells. Not for a sheet added to a shared workbook (Report Filter Pages): one pivot per file.
  asPivotTable = true,
  // Formats set on individual cells (Home › Number, utils/cellFormats.js) — they win over the value
  // field's format, like the live grid.
  cellFormats = null,
}) {
  sheetName = sheetName || pivotOptions.name?.trim() || "Pivot";
  // When Excel can rebuild these numbers from detail rows, the sheet becomes a PivotTable; the cells
  // drawn below are then only what shows until Excel refreshes it on open (Protected View), so they
  // stay plain — no merges, fills or conditional formats, which would get in the pivot's way.
  // Cell formats can't be carried by the generated PivotTable (it formats whole value fields), so a
  // pivot with any stays plain values that keep them.
  const pivotSource = !asPivotTable || targetWorkbook
    ? null
    : hasCellFormats(cellFormats)
      ? { ok: false, reason: "cells have their own number formats, which are kept this way" }
      : buildPivotSource({ result, rows, columns, rowFieldLabels, columnFieldLabels, valueFields, filters, filterFieldLabels, drilldown, drillSlot, columnsDrillSlot });
  const plain = Boolean(pivotSource?.ok);
  // As a PivotTable, the filter area holds just the page fields Excel will draw there (a field at
  // (All) has none — see buildPivotSource), so no stale caption is left beside them.
  if (plain) filters = filters.filter((f) => !rows.includes(f.field) && !columns.includes(f.field) && filterValueSummary(f) !== "(All)");
  const { rowHeaders, columnHeaders, cells } = result;
  // MAX across every row — a drilldown's rows have different depths.
  const labelDepth = rowHeaders.reduce((max, h) => Math.max(max, h.labels.length), 0);
  const rowFieldLayouts = Array.from({ length: labelDepth }, (_, level) => fieldLayouts.find((f) => f.field === rows[level]) ?? null);
  const rowHeaderGroups = drilldown
    ? [{ start: 0, end: Math.max(0, labelDepth - 1) }]
    : buildRowHeaderGroups(labelDepth, rowFieldLayouts.map((fl) => fl?.compact ?? layout === "Compact"));
  // One sheet column per row-header GROUP (a compact-merged group shares one column), not per level.
  const rowDepth = rowHeaderGroups.length;
  // Same floor-of-1 fallback PivotGrid.jsx uses for its own on-screen rendering (see that file's own
  // WHY comment) — with no Value fields at all, Columns can legitimately come back with ZERO
  // headers, but the row-field title row below still needs ONE pass of the
  // `for (let level = 0; level < columnDepth; level++)` loop to render at all. Missing here meant
  // the exported .xlsx silently dropped the row headers' own title row for a Values-less pivot,
  // even though the live grid correctly showed it (caught 2026-10-06, full codebase review).
  const allColumnHeaderRows = buildColumnHeaderRows(columnHeaders, columnsDrillSlot, measuresIndexOf(columns));
  // One value field with Columns fields: no row repeating its name under every column — like the
  // live grid and Excel, the name goes in the top-left corner instead (2026-10-21).
  const labelLevels = columnHeaders[0]?.labels.length ?? 0;
  const singleValueRowIndex = valueFields.length === 1 && columnsDrillSlot < 0 && labelLevels > 1 && measuresIndexOf(columns) === labelLevels - 1 ? labelLevels - 1 : -1;
  const columnHeaderRows = singleValueRowIndex >= 0 ? allColumnHeaderRows.filter((_, r) => r !== singleValueRowIndex) : allColumnHeaderRows;
  const singleValueCaption = singleValueRowIndex >= 0 ? (columnHeaders.find((c) => !c.isTotal)?.labels[singleValueRowIndex] ?? "") : null;
  const columnDepth = columnHeaders.length > 0 ? columnHeaderRows.length : rowHeaders.length > 0 ? 1 : 0;
  // Where (if either) Σ Values sits — rowFieldLabels/columnFieldLabels already carry "Σ Values" at
  // that same position (PivotBuilder.jsx's fieldLabels helper maps the sentinel to that label), so
  // no separate "append Σ Values" step is needed here anymore.
  const rowsMeasuresIndex = measuresIndexOf(rows);
  const columnsMeasuresIndex = measuresIndexOf(columns);
  const measuresInRows = rowsMeasuresIndex >= 0;
  const captionToIndex = valueFields.length
    ? buildCaptionToIndex(measuresInRows ? rowHeaders : columnHeaders, measuresInRows ? rowsMeasuresIndex : columnsMeasuresIndex, valueFields.length)
    : new Map();

  // Report Filter area: each field is a caption cell + a selection cell, placed by filterAreaSlot
  // (a blank column between the pairs, like Excel), then one blank row before the table.
  const filterSlots = filters.map((f, i) => filterAreaSlot(i, filters.length, pivotOptions));
  const filterAreaRows = filters.length > 0 ? Math.max(...filterSlots.map((s) => s.row)) + 2 : 0;

  const workbook = targetWorkbook ?? new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(
    safeSheetName(workbook, sheetName),
    plain ? {} : { views: [{ state: "frozen", xSplit: rowDepth, ySplit: filterAreaRows + columnDepth + (columnFieldLabels.length > 0 ? 1 : 0) }] },
  );
  const mergeCells = (...range) => {
    if (!plain) sheet.mergeCells(...range);
  };

  filters.forEach((filter, i) => {
    const { column, row } = filterSlots[i];
    const caption = sheet.getRow(row + 1).getCell(column * 3 + 1);
    caption.value = filterFieldLabels[filter.field] ?? filter.field;
    caption.font = { bold: true };
    caption.fill = HEADER_FILL;
    sheet.getRow(row + 1).getCell(column * 3 + 2).value = filterValueSummary(filter);
  });

  const showCaptions = pivotOptions.showFieldCaptions !== false;
  const totalCols = rowDepth + columnHeaders.length;
  let currentRow = 1 + filterAreaRows;

  // Row 1: the crossed field label (e.g. "Doctor.Doctor Name"), merged across the data columns —
  // mirrors PivotGrid's own muted pill row above the column headers.
  if (columnFieldLabels.length > 0) {
    const labelRow = sheet.getRow(currentRow);
    if (rowDepth + 1 <= totalCols) mergeCells(currentRow, rowDepth + 1, currentRow, totalCols);
    const cell = labelRow.getCell(rowDepth + 1);
    // The hidden value row takes its "Σ Values" caption with it.
    const shownColumnLabels = singleValueCaption !== null ? columnFieldLabels.filter((_, i) => i !== columnsMeasuresIndex) : columnFieldLabels;
    cell.value = showCaptions ? shownColumnLabels.join(" › ") : "";
    if (singleValueCaption && rowDepth > 0) {
      const corner = labelRow.getCell(1);
      corner.value = singleValueCaption;
      corner.font = { bold: true };
    }
    cell.font = { italic: true, color: { argb: "FF64748B" } };
    cell.alignment = { horizontal: "center" };
    currentRow += 1;
  }

  // One row per column-header level, with the row-header field labels occupying the top-left
  // corner (merged down through every column-header level, same as PivotGrid's rowSpan).
  const columnHeaderStartRow = currentRow;
  for (let level = 0; level < columnDepth; level++) {
    const row = sheet.getRow(currentRow);
    if (level === 0) {
      rowHeaderGroups.forEach((group, gi) => {
        const cell = row.getCell(gi + 1);
        // A compact-merged column is Excel's literal "Row Labels"; a standalone one is its field.
        // Blank with "Display field captions and filter drop downs" off (PivotTable Options > Display).
        cell.value = !showCaptions ? "" : group.start < group.end ? "Row Labels" : rowFieldLayouts[group.start]?.customName || rowFieldLabels[group.start] || "";
        cell.font = { bold: true };
        cell.fill = HEADER_FILL;
        if (columnDepth > 1) mergeCells(columnHeaderStartRow, gi + 1, columnHeaderStartRow + columnDepth - 1, gi + 1);
      });
    }
    // A Grand Total column shows its caption once, in the top header row, like Excel.
    const levelLabels = columnHeaders.map((c) => (c.captionLevel === -1 && c.caption ? (level === 0 ? c.caption : "") : c.labels[level]));
    // The same spans as the live grid (buildColumnHeaderRows: an outer item over its children, a
    // "2024 Total" caption, blanks under it), since 2026-10-15 for ordinary Columns fields too.
    const runs = columnHeaderRows[level]?.map((cell) => ({ start: cell.sourceIndex, end: cell.sourceIndex + cell.span, label: cell.text })) ?? mergeRuns(levelLabels);
    for (const run of runs) {
      const startCol = rowDepth + run.start + 1;
      const endCol = rowDepth + run.end;
      const cell = row.getCell(startCol);
      cell.value = run.label;
      cell.font = { bold: true };
      cell.fill = HEADER_FILL;
      // An outer header spanning its children is left-aligned (Excel's Compact layout) — centered with
      // "Merge and center cells with labels"; a single column stays centered.
      cell.alignment = { horizontal: endCol > startCol && !pivotOptions.mergeLabels ? "left" : "center" };
      if (endCol > startCol) mergeCells(currentRow, startCol, currentRow, endCol);
    }
    currentRow += 1;
  }

  // Row headers + values, one sheet row per DISPLAY row (see buildDisplayRows) — the same rows the
  // live grid draws: compact-form item rows (indented with Excel's own cell indent), "ER Total"/
  // "Grand Total" captions, each label shown once (Excel doesn't merge label cells by default).
  const displayRows = buildDisplayRows(rowHeaders, { rowHeaderGroups, rowFieldLayouts, repeatItemLabels, measuresLevel: rowsMeasuresIndex, outline: layout !== "Tabular", drilldown, drilldownSubtotals, drillSlot });
  const firstDataRow = currentRow;

  // PivotTable Options › Printing (2026-10-20), applied where printing happens — the exported sheet:
  // "Set print titles" repeats the column header rows at the top of every printed page, and "Repeat
  // row labels on each printed page" the row label columns at the left of every page.
  if (pivotOptions.printTitles && firstDataRow > 1 + filterAreaRows) {
    sheet.pageSetup.printTitlesRow = `${1 + filterAreaRows}:${firstDataRow - 1}`;
  }
  if (pivotOptions.repeatRowLabelsOnPrint !== false && rowDepth > 0) {
    sheet.pageSetup.printTitlesColumn = `A:${sheet.getColumn(rowDepth).letter}`;
  }
  // Alt Text: Excel stores it on the PivotTable; here it goes into the file's own Title/Subject.
  if (pivotOptions.altTextTitle) workbook.title = pivotOptions.altTextTitle;
  if (pivotOptions.altTextDescription) workbook.subject = pivotOptions.altTextDescription;
  // "Merge and center cells with labels": a real merged cell per item (computeLabelSpans). A merge
  // never crosses a blank line, so its rows are consecutive sheet rows.
  const labelSpans = pivotOptions.mergeLabels ? computeLabelSpans(displayRows, rowHeaderGroups, { drillSlot, drilldown }) : null;
  // Sheet row of every display row that shows a result row's values: [excelRow, rowIndex, kind].
  const placedRows = [];
  displayRows.forEach((displayRow, displayIndex) => {
    const excelRow = currentRow;
    const isTotalRow = displayRow.kind === "subtotal" || displayRow.kind === "grand" || (displayRow.kind === "header" && displayRow.sourceIndex !== null);
    rowHeaderGroups.forEach((_, gi) => {
      const span = labelSpans?.[displayIndex][gi] ?? -1;
      if (span === 0) return; // inside the merged cell above
      const cell = sheet.getRow(excelRow).getCell(gi + 1);
      const lines = displayRow.groupCells[gi]?.lines ?? [];
      cell.value = lines.map((l) => l.text).join(" › ");
      if (lines[0]?.indent) cell.alignment = { indent: lines[0].indent };
      // An outer item's label (span >= 1) is merged and centered; everything else keeps its alignment.
      if (span >= 1) cell.alignment = { horizontal: "center", vertical: "middle" };
      if (span > 1) mergeCells(excelRow, gi + 1, excelRow + span - 1, gi + 1);
      cell.font = { bold: isTotalRow || displayRow.kind === "header" };
      if (isTotalRow) cell.fill = SUBTOTAL_FILL;
    });
    if (displayRow.sourceIndex !== null) placedRows.push([excelRow, displayRow.sourceIndex, displayRow.kind]);
    currentRow += displayRow.blankLineAfter ? 2 : 1;
  });

  // Data cells — same measure-lookup rule as PivotGrid: whichever axis actually carries Σ Values
  // determines which PivotValueField a given cell belongs to, read directly off that header's own
  // caption at the sentinel's position (not a fixed stride — see utils/measuresPlacement.js).
  placedRows.forEach(([excelRow, rowIndex]) => {
    const rowHeader = rowHeaders[rowIndex];
    columnHeaders.forEach((colHeader, colIndex) => {
      const measureIndex = measuresInRows
        ? measureIndexOf(rowHeader, rowsMeasuresIndex, captionToIndex)
        : measureIndexOf(colHeader, columnsMeasuresIndex, captionToIndex);
      const value = cells[rowIndex][colIndex];
      const isTotalCell = rowHeader.isTotal || colHeader.isTotal;
      const cell = sheet.getRow(excelRow).getCell(rowDepth + colIndex + 1);
      // An error cell is a real Excel #DIV/0! (or the "For error values show" text); an empty one
      // stays empty unless "For empty cells show" has text.
      if (isErrorValue(value)) {
        cell.value = pivotOptions.showErrorText ? pivotOptions.errorText : { error: ERROR_TOKEN };
      } else if (value === null || value === undefined) {
        cell.value = pivotOptions.showEmptyText && pivotOptions.emptyText ? pivotOptions.emptyText : null;
      } else {
        cell.value = value;
      }
      cell.numFmt = excelNumberFormat(cellFormats?.[cellFormatKey(rowHeader, colHeader)] ?? formatForColumn(measureIndex, valueFields));
      cell.font = { bold: isTotalCell };
      cell.alignment = { horizontal: "center" };
      if (isTotalCell) cell.fill = SUBTOTAL_FILL;
    });
  });

  // Plain cells under a PivotTable: Excel applies the pivot's own style once it refreshes.
  if (plain) {
    sheet.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => (cell.style = {})));
  }

  // Native Excel conditional formatting, scoped to each measure's own detail cells (excluding
  // Total/Grand-Total rows or columns — same exclusion detailRangeForColumn/Row already applies
  // for the live grid's min/max, done here as actual excluded RANGES instead, so Excel computes
  // min/max live over just the real data rather than a snapshot of it).
  if (plain) {
    // (none — see above)
  } else if (measuresInRows) {
    placedRows.forEach(([excelRow, rowIndex]) => {
      const rowHeader = rowHeaders[rowIndex];
      if (rowHeader.isTotal) return;
      const measureIndex = measureIndexOf(rowHeader, rowsMeasuresIndex, captionToIndex);
      const valueField = measureIndex !== null ? valueFields[measureIndex] : null;
      const rule = excelConditionalFormatRule(valueField?.conditionalFormat);
      if (!rule) return;
      const totalColIndexes = new Set(columnHeaders.map((c, i) => (c.isTotal ? i : null)).filter((i) => i !== null));
      const ref = detailRuns(columnHeaders.length, totalColIndexes)
        .map((run) => `${sheet.getColumn(rowDepth + run.start + 1).letter}${excelRow}:${sheet.getColumn(rowDepth + run.end).letter}${excelRow}`)
        .join(" ");
      if (ref) sheet.addConditionalFormatting({ ref, rules: [rule] });
    });
  } else {
    columnHeaders.forEach((colHeader, colIndex) => {
      if (colHeader.isTotal) return;
      const measureIndex = measureIndexOf(colHeader, columnsMeasuresIndex, captionToIndex);
      const valueField = measureIndex !== null ? valueFields[measureIndex] : null;
      const rule = excelConditionalFormatRule(valueField?.conditionalFormat);
      if (!rule) return;
      const excelCol = sheet.getColumn(rowDepth + colIndex + 1).letter;
      // Sheet rows that hold DETAIL values (totals, blank-valued item rows and blank lines excluded).
      const sheetRowCount = currentRow - firstDataRow;
      const detailOffsets = new Set(placedRows.filter(([, rowIndex]) => !rowHeaders[rowIndex].isTotal).map(([excelRow]) => excelRow - firstDataRow));
      const excluded = new Set(Array.from({ length: sheetRowCount }, (_, i) => i).filter((i) => !detailOffsets.has(i)));
      const ref = detailRuns(sheetRowCount, excluded)
        .map((run) => `${excelCol}${firstDataRow + run.start}:${excelCol}${firstDataRow + run.end - 1}`)
        .join(" ");
      if (ref) sheet.addConditionalFormatting({ ref, rules: [rule] });
    });
  }

  // Reasonable auto-width per column instead of Excel's cramped default — based on the widest
  // header/label text actually placed in that column, capped so one long outlier doesn't blow out
  // the whole sheet.
  // (The report filter area can reach past the table's own columns.)
  const filterAreaCols = filters.length > 0 ? Math.max(...filterSlots.map((s) => s.column)) * 3 + 2 : 0;
  for (let col = 1; col <= Math.max(totalCols, filterAreaCols); col++) {
    let maxLen = 8;
    sheet.getColumn(col).eachCell({ includeEmpty: false }, (cell) => {
      const text = cell.value?.error ?? cell.value?.toString() ?? "";
      if (text.length > maxLen) maxLen = text.length;
    });
    sheet.getColumn(col).width = Math.min(maxLen + 2, 40);
  }

  if (targetWorkbook) return workbook;
  if (!plain) {
    await downloadWorkbook(workbook, filename);
    return { workbook, pivotTable: false, reason: pivotSource?.reason ?? null };
  }

  // The PivotTable's source: one row per Rows × Columns item combination (buildPivotSource), on its
  // own sheet after the pivot, measure columns in each value field's number format.
  const dataSheet = workbook.addWorksheet(safeSheetName(workbook, "Data"));
  dataSheet.addRow(pivotSource.fields.map((f) => f.name)).font = { bold: true };
  pivotSource.data.forEach((record) => dataSheet.addRow(record.map((v) => (v === "" ? null : v))));
  const numFmtCodes = [];
  pivotSource.fields.forEach((f, i) => {
    const column = dataSheet.getColumn(i + 1);
    column.width = Math.min(Math.max(f.name.length, ...pivotSource.data.slice(0, 200).map((r) => String(r[i] ?? "").length)) + 2, 40);
    if (f.kind !== "data") return;
    const code = excelNumberFormat(formatForColumn(f.measureIndex, valueFields));
    numFmtCodes.push(code);
    column.numFmt = code;
  });
  const lastDataCol = dataSheet.getColumn(pivotSource.fields.length).letter;
  const pivotTop = 1 + filterAreaRows;
  const buffer = await workbook.xlsx.writeBuffer();
  const withPivot = await addPivotTableToXlsx(buffer, {
    source: pivotSource,
    name: pivotOptions.name?.trim() || "PivotTable1",
    pivotSheet: sheet.name,
    dataSheet: dataSheet.name,
    dataRef: `A1:${lastDataCol}${pivotSource.data.length + 1}`,
    location: `A${pivotTop}:${sheet.getColumn(Math.max(totalCols, 1)).letter}${Math.max(currentRow - 1, pivotTop + 1)}`,
    layout,
    subtotalsAtBottom: drilldownSubtotals === "Bottom",
    numFmtCodes,
    pageOverThenDown: pivotOptions.filterAreaOrder === "OverThenDown",
    pageWrap: pivotOptions.filterFieldsPerColumn > 0 ? pivotOptions.filterFieldsPerColumn : 0,
  });
  downloadBlob(withPivot, filename);
  return { workbook, pivotTable: true, verified: pivotSource.verified };
}

// Excel sheet names: at most 31 characters, none of \ / ? * [ ] :, and unique in the workbook. An
// apostrophe becomes ’ — the print-titles name ExcelJS writes quotes the sheet name as 'name'!…
// without escaping, so "Children's" made Excel report a damaged file (found in review 2026-10-20).
function safeSheetName(workbook, name) {
  const base = String(name ?? "").replace(/[\\/?*[\]:]/g, " ").replace(/'/g, "’").trim().slice(0, 31) || "Sheet";
  let candidate = base;
  for (let n = 2; workbook.getWorksheet(candidate); n++) {
    const suffix = ` (${n})`;
    candidate = base.slice(0, 31 - suffix.length) + suffix;
  }
  return candidate;
}

export function createWorkbook() {
  return new ExcelJS.Workbook();
}

// Saves the workbook as "<filename>.xlsx". The blob URL is released only after a delay — revoking
// it right after click() let Chrome start the download from a URL that was already gone, and the
// file landed with the blob's GUID as its name and no extension (caught live 2026-10-20: Export
// from Snapshot History saved "5ea6bf54-…" instead of "<pivot>_v1.xlsx"). Characters Windows
// doesn't allow in a file name are replaced too, so the name is never rejected.
export async function downloadWorkbook(workbook, filename) {
  downloadBlob(await workbook.xlsx.writeBuffer(), filename);
}

function downloadBlob(buffer, filename) {
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const safeName = String(filename ?? "").replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim() || "Pivot";
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${safeName}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

// Show Details → Excel (2026-10-20): what Excel's own Show Details produces — the underlying rows on
// a new sheet as a formatted Excel Table (filter buttons, banded rows). Numeric text from the cube
// becomes a real number so it sums/sorts in Excel.
export async function exportDetailsToExcel({ columns, rows, filename, sheetName = "Details" }) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(safeSheetName(workbook, sheetName));

  // Table column names must be unique and non-empty.
  const seen = new Map();
  const names = columns.map((c, i) => {
    const base = String(c ?? "").trim() || `Column${i + 1}`;
    const n = (seen.get(base) ?? 0) + 1;
    seen.set(base, n);
    return n === 1 ? base : `${base} (${n})`;
  });
  // Only a plain decimal becomes a number — not an ID with leading zeros ("000123"), one too long to
  // keep exactly (16+ digits), or hex/exponent text that Number() would also accept.
  const toCell = (v) => {
    if (typeof v !== "string") return v ?? null;
    const t = v.trim();
    const plainDecimal = /^-?(0|[1-9]\d*)(\.\d+)?$/.test(t) && t.replace(/[-.]/g, "").length <= 15;
    return plainDecimal ? Number(t) : v;
  };

  sheet.addTable({
    name: "Details",
    ref: "A1",
    headerRow: true,
    style: { theme: "TableStyleMedium2", showRowStripes: true },
    columns: names.map((name) => ({ name, filterButton: true })),
    rows: rows.length > 0 ? rows.map((r) => names.map((_, i) => toCell(r[i]))) : [names.map(() => null)],
  });

  names.forEach((name, i) => {
    const longest = Math.max(name.length, ...rows.slice(0, 200).map((r) => String(r[i] ?? "").length));
    sheet.getColumn(i + 1).width = Math.min(longest + 3, 50);
  });

  await downloadWorkbook(workbook, filename);
}
