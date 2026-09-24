import ExcelJS from "exceljs";
import { formatForColumn } from "./numberFormat";
import { computeVisibleLevels } from "./pivotGridLayout";

const SUBTOTAL_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF6E9D8" } }; // ~bg-gold/10
const HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } }; // slate-100

// Same clamp/decimal handling as numberFormat.js's formatNumber, but producing an EXCEL number
// format CODE instead of a display string — the cell keeps its real numeric value (sortable,
// summable in Excel) with Excel doing the formatting, instead of exporting pre-formatted text.
//
// Percentage is the one non-obvious case: our values are already on a 0–100 scale (38.43, not
// 0.3843 — see numberFormat.js), but Excel's real "0.00%" format code MULTIPLIES the underlying
// number by 100 for display. Using that would double-apply the scaling (38.43 -> "3843.00%").
// Appending a literal quoted "%" character instead (not Excel's percent operator) shows the
// number exactly as-is with a % suffix, matching this app's own convention everywhere else.
function excelNumberFormat(format) {
  const type = format?.type ?? "General";
  const decimals = Math.min(Math.max(format?.decimalPlaces ?? 2, 0), 10);
  const decimalPart = decimals > 0 ? "." + "0".repeat(decimals) : "";
  if (type === "Number") return `#,##0${decimalPart}`;
  if (type === "Currency") return `"${(format?.currencySymbol ?? "EGP").replace(/"/g, "")}" #,##0${decimalPart}`;
  if (type === "Percentage") return `#,##0${decimalPart}"%"`;
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
      color: [{ argb: "FFF8696B" }, { argb: "FFFFEB84" }, { argb: "FF63BE7B" }],
    };
  }
  if (format.type === "DataBar") {
    return {
      type: "dataBar",
      cfvo: [{ type: "min" }, { type: "max" }],
      color: { argb: "FF638EC6" },
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
  rowFieldLabels = [],
  columnFieldLabels = [],
  valueFields = [],
  valuesPlacement = "Columns",
  filename = "Pivot",
}) {
  const { rowHeaders, columnHeaders, cells } = result;
  const rowDepth = rowHeaders[0]?.labels.length ?? 0;
  const columnDepth = columnHeaders[0]?.labels.length ?? 0;
  const rowFieldLabelsWithMeasure = valuesPlacement === "Rows" ? [...rowFieldLabels, "Σ Values"] : rowFieldLabels;

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
      for (let rl = 0; rl < rowDepth; rl++) {
        const cell = row.getCell(rl + 1);
        cell.value = rowFieldLabelsWithMeasure[rl] ?? "";
        cell.font = { bold: true };
        cell.fill = HEADER_FILL;
        if (columnDepth > 1) sheet.mergeCells(columnHeaderStartRow, rl + 1, columnHeaderStartRow + columnDepth - 1, rl + 1);
      }
    }
    const levelLabels = columnHeaders.map((c) => c.labels[level]);
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

  // Row headers: each level vertically merges across consecutive rows sharing the same label —
  // same "Tabular Form" suppression as computeVisibleLevels, expressed as real merged cells
  // instead of blank ones (there's no live layout toggle here, so this always uses that shape
  // regardless of the on-screen Tabular/Compact choice — Compact's single indented column is a
  // screen/print concession a spreadsheet doesn't need).
  const firstDataRow = currentRow;
  let previousLabels = null;
  const pendingMergeStart = new Array(rowDepth).fill(null);
  rowHeaders.forEach((rowHeader, rowIndex) => {
    const excelRow = firstDataRow + rowIndex;
    const visible = computeVisibleLevels(rowHeader.labels, rowHeader.isTotal ? null : previousLabels);
    for (let level = 0; level < rowDepth; level++) {
      const cell = sheet.getRow(excelRow).getCell(level + 1);
      if (visible[level]) {
        if (pendingMergeStart[level] !== null && excelRow - 1 > pendingMergeStart[level]) {
          sheet.mergeCells(pendingMergeStart[level], level + 1, excelRow - 1, level + 1);
        }
        cell.value = rowHeader.labels[level] ?? "";
        pendingMergeStart[level] = excelRow;
      }
      cell.font = { bold: rowHeader.isTotal };
      if (rowHeader.isTotal) cell.fill = SUBTOTAL_FILL;
    }
    previousLabels = rowHeader.isTotal ? null : rowHeader.labels;
  });
  // Close out any merge run still open after the last row.
  for (let level = 0; level < rowDepth; level++) {
    const lastRow = firstDataRow + rowHeaders.length - 1;
    if (pendingMergeStart[level] !== null && lastRow > pendingMergeStart[level]) {
      sheet.mergeCells(pendingMergeStart[level], level + 1, lastRow, level + 1);
    }
  }

  // Data cells — same measure-lookup rule as PivotGrid ("measures are the fastest-varying
  // factor"): whichever axis actually carries the crossjoin (Columns by default, or Rows when
  // ValuesPlacement.Rows) determines which PivotValueField a given cell belongs to.
  rowHeaders.forEach((rowHeader, rowIndex) => {
    const excelRow = firstDataRow + rowIndex;
    columnHeaders.forEach((colHeader, colIndex) => {
      const measureIndex = valuesPlacement === "Rows" ? rowIndex : colIndex;
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
  if (valuesPlacement === "Rows") {
    rowHeaders.forEach((rowHeader, rowIndex) => {
      if (rowHeader.isTotal) return;
      const measureIndex = rowIndex;
      const valueField = valueFields.length ? valueFields[measureIndex % valueFields.length] : null;
      const rule = excelConditionalFormatRule(valueField?.conditionalFormat);
      if (!rule) return;
      const excelRow = firstDataRow + rowIndex;
      const totalColIndexes = new Set(columnHeaders.map((c, i) => (c.isTotal ? i : null)).filter((i) => i !== null));
      const ref = detailRuns(columnHeaders.length, totalColIndexes)
        .map((run) => `${sheet.getColumn(rowDepth + run.start + 1).letter}${excelRow}:${sheet.getColumn(rowDepth + run.end).letter}${excelRow}`)
        .join(" ");
      if (ref) sheet.addConditionalFormatting({ ref, rules: [rule] });
    });
  } else {
    columnHeaders.forEach((colHeader, colIndex) => {
      if (colHeader.isTotal) return;
      const measureIndex = colIndex;
      const valueField = valueFields.length ? valueFields[measureIndex % valueFields.length] : null;
      const rule = excelConditionalFormatRule(valueField?.conditionalFormat);
      if (!rule) return;
      const excelCol = sheet.getColumn(rowDepth + colIndex + 1).letter;
      const totalRowIndexes = new Set(rowHeaders.map((r, i) => (r.isTotal ? i : null)).filter((i) => i !== null));
      const ref = detailRuns(rowHeaders.length, totalRowIndexes)
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
