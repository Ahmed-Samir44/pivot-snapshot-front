import { formatNumber } from "./numberFormat";
import { isErrorValue } from "./errorCells";

// Excel's PivotTable Options > Layout & Format tab (2026-10-18, against a screenshot of the real
// dialog). Display-only, kept beside `request` like fieldLayouts so a change never re-queries the
// cube. Mirrors the backend's PivotTableOptions record exactly.
//   mergeLabels: "Merge and center cells with labels".
//   showErrorText/errorText: "For error values show" — off shows Excel's own #DIV/0!.
//   showEmptyText/emptyText: "For empty cells show" — on with "" (Excel's default) is a blank cell.
//   filterAreaOrder/filterFieldsPerColumn: "Display fields in report filter area" and "Report
//     filter fields per column" (0 = all in one column/row, Excel's default).
export const DEFAULT_PIVOT_OPTIONS = {
  name: "PivotTable1",
  mergeLabels: false,
  showErrorText: false,
  errorText: "",
  showEmptyText: true,
  emptyText: "",
  filterAreaOrder: "DownThenOver",
  filterFieldsPerColumn: 0,
};

export const ERROR_TOKEN = "#DIV/0!";

// What a VALUE cell shows: the empty/error text for a blank or error cell, else the formatted
// number. Mirrors PivotTableOptions.TextFor + NumberFormatter.Format.
export function cellText(value, format, options = DEFAULT_PIVOT_OPTIONS) {
  if (value === null || value === undefined) return options.showEmptyText ? options.emptyText ?? "" : "";
  if (isErrorValue(value)) return options.showErrorText ? options.errorText ?? "" : ERROR_TOKEN;
  return formatNumber(value, format);
}

// Down, Then Over: N fields down each column, then the next column; Over, Then Down: N fields across
// each row, then the next row. N = 0 means no limit. Mirrors HtmlSnapshotRenderer.FilterAreaSlot.
export function filterAreaSlot(index, count, options = DEFAULT_PIVOT_OPTIONS) {
  const perLine = options.filterFieldsPerColumn > 0 ? options.filterFieldsPerColumn : Math.max(count, 1);
  return options.filterAreaOrder === "OverThenDown"
    ? { column: index % perLine, row: Math.floor(index / perLine) }
    : { column: Math.floor(index / perLine), row: index % perLine };
}

// Excel's caption for a Report Filter's current selection: "(All)" when nothing narrows the field,
// the member itself when exactly one is picked, "(Multiple Items)" otherwise (label/Top N/range modes
// included). Mirrors HtmlSnapshotRenderer.FilterValueSummary.
export function filterValueSummary(filter) {
  if (filter.mode === "Members" || filter.mode === "Exclude" || !filter.mode) {
    if (filter.includedMembers.length === 0) return "(All)";
    if (filter.includedMembers.length === 1) return filter.includedMembers[0] === "" ? "(blank)" : filter.includedMembers[0];
    return "(Multiple Items)";
  }
  if (filter.mode === "LabelContains" || filter.mode === "LabelBeginsWith" || filter.mode === "LabelEndsWith") {
    return filter.labelText ? "(Multiple Items)" : "(All)";
  }
  return "(Multiple Items)";
}
