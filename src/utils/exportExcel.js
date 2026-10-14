import ExcelJS from "exceljs";
import { formatForColumn } from "./numberFormat";
import { buildRowHeaderGroups, buildDisplayRows } from "./pivotGridLayout";
import { measuresIndexOf, buildCaptionToIndex, measureIndexOf } from "./measuresPlacement";
import { DEFAULT_MIN_COLOR, DEFAULT_MID_COLOR, DEFAULT_MAX_COLOR, DEFAULT_BAR_COLOR } from "./conditionalFormat";

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
  if (type === "Number") return `#,##0${decimalPart}`;
  if (type === "Currency") return `${currencyPrefix}#,##0${decimalPart}`;
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
}) {
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
  const columnDepth = columnHeaders[0]?.labels.length ?? (rowHeaders.length > 0 ? 1 : 0);
  // Where (if either) Σ Values sits — rowFieldLabels/columnFieldLabels already carry "Σ Values" at
  // that same position (PivotBuilder.jsx's fieldLabels helper maps the sentinel to that label), so
  // no separate "append Σ Values" step is needed here anymore.
  const rowsMeasuresIndex = measuresIndexOf(rows);
  const columnsMeasuresIndex = measuresIndexOf(columns);
  const measuresInRows = rowsMeasuresIndex >= 0;
  const captionToIndex = valueFields.length
    ? buildCaptionToIndex(measuresInRows ? rowHeaders : columnHeaders, measuresInRows ? rowsMeasuresIndex : columnsMeasuresIndex, valueFields.length)
    : new Map();

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Pivot", { views: [{ state: "frozen", xSplit: rowDepth, ySplit: columnDepth + (columnFieldLabels.length > 0 ? 1 : 0) }] });

  const totalCols = rowDepth + columnHeaders.length;
  let currentRow = 1;

  // Row 1: the crossed field label (e.g. "Doctor.Doctor Name"), merged across the data columns —
  // mirrors PivotGrid's own muted pill row above the column headers.
  if (columnFieldLabels.length > 0) {
    const labelRow = sheet.getRow(currentRow);
    if (rowDepth + 1 <= totalCols) sheet.mergeCells(currentRow, rowDepth + 1, currentRow, totalCols);
    const cell = labelRow.getCell(rowDepth + 1);
    cell.value = columnFieldLabels.join(" › ");
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
        cell.value = group.start < group.end ? "Row Labels" : rowFieldLayouts[group.start]?.customName || rowFieldLabels[group.start] || "";
        cell.font = { bold: true };
        cell.fill = HEADER_FILL;
        if (columnDepth > 1) sheet.mergeCells(columnHeaderStartRow, gi + 1, columnHeaderStartRow + columnDepth - 1, gi + 1);
      });
    }
    // A Grand Total column shows its caption once, in the top header row, like Excel.
    const levelLabels = columnHeaders.map((c) => (c.captionLevel === -1 && c.caption ? (level === 0 ? c.caption : "") : c.labels[level]));
    for (const run of mergeRuns(levelLabels)) {
      const startCol = rowDepth + run.start + 1;
      const endCol = rowDepth + run.end;
      const cell = row.getCell(startCol);
      cell.value = run.label;
      cell.font = { bold: true };
      cell.fill = HEADER_FILL;
      cell.alignment = { horizontal: "center" };
      if (endCol > startCol) sheet.mergeCells(currentRow, startCol, currentRow, endCol);
    }
    currentRow += 1;
  }

  // Row headers + values, one sheet row per DISPLAY row (see buildDisplayRows) — the same rows the
  // live grid draws: compact-form item rows (indented with Excel's own cell indent), "ER Total"/
  // "Grand Total" captions, each label shown once (Excel doesn't merge label cells by default).
  const displayRows = buildDisplayRows(rowHeaders, { rowHeaderGroups, rowFieldLayouts, repeatItemLabels, measuresLevel: rowsMeasuresIndex, outline: layout !== "Tabular", drilldown, drilldownSubtotals });
  const firstDataRow = currentRow;
  // Sheet row of every display row that shows a result row's values: [excelRow, rowIndex, kind].
  const placedRows = [];
  displayRows.forEach((displayRow) => {
    const excelRow = currentRow;
    const isTotalRow = displayRow.kind === "subtotal" || displayRow.kind === "grand" || (displayRow.kind === "header" && displayRow.sourceIndex !== null);
    rowHeaderGroups.forEach((_, gi) => {
      const cell = sheet.getRow(excelRow).getCell(gi + 1);
      const lines = displayRow.groupCells[gi]?.lines ?? [];
      cell.value = lines.map((l) => l.text).join(" › ");
      if (lines[0]?.indent) cell.alignment = { indent: lines[0].indent };
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
      cell.value = value ?? null;
      cell.numFmt = excelNumberFormat(formatForColumn(measureIndex, valueFields));
      cell.font = { bold: isTotalCell };
      cell.alignment = { horizontal: "center" };
      if (isTotalCell) cell.fill = SUBTOTAL_FILL;
    });
  });

  // Native Excel conditional formatting, scoped to each measure's own detail cells (excluding
  // Total/Grand-Total rows or columns — same exclusion detailRangeForColumn/Row already applies
  // for the live grid's min/max, done here as actual excluded RANGES instead, so Excel computes
  // min/max live over just the real data rather than a snapshot of it).
  if (measuresInRows) {
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
  for (let col = 1; col <= totalCols; col++) {
    let maxLen = 8;
    sheet.getColumn(col).eachCell({ includeEmpty: false }, (cell) => {
      const text = cell.value?.toString() ?? "";
      if (text.length > maxLen) maxLen = text.length;
    });
    sheet.getColumn(col).width = Math.min(maxLen + 2, 40);
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${filename}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}
