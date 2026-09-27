import { useMemo, useState } from "react";
import { ChevronDown, ChevronRight, Search } from "lucide-react";
import { formatNumber, formatForColumn } from "../../utils/numberFormat";
import { conditionalStyleFor, detailRangeForColumn, detailRangeForRow, iconColorFor } from "../../utils/conditionalFormat";
import { computeGroups, computeCompactVisibility } from "../../utils/pivotGridLayout";

// A pill-shaped header badge (dimension/measure name + a decorative caret), echoing the rounded
// filter-pill style used elsewhere in the product's screens, instead of a traditional shaded
// spreadsheet header cell.
function HeaderPill({ children, muted = false }) {
  return (
    <span
      className={`inline-flex items-center gap-1.5 whitespace-nowrap rounded-full border px-4 py-2 text-sm font-semibold shadow-sm ${
        muted ? "border-transparent bg-transparent text-muted shadow-none" : "border-slate-200 bg-white text-ink"
      }`}
    >
      {children}
      {!muted && <ChevronDown className="h-3.5 w-3.5 text-muted" />}
    </span>
  );
}

export default function PivotGrid({
  result,
  rowFieldLabels = [],
  columnFieldLabels = [],
  valueFields = [],
  layout = "Tabular",
  valuesPlacement = "Columns",
}) {
  const [collapsed, setCollapsed] = useState(() => new Set());
  const [searchText, setSearchText] = useState("");

  const groups = useMemo(() => (result ? computeGroups(result.rowHeaders) : []), [result]);
  // Excel's own literal label for the pseudo-field, appended as one more row-header level when
  // measures are crossjoined into Rows instead of Columns — mirrors HtmlSnapshotRenderer exactly
  // (row.labels themselves already carry the measure name automatically, no change needed there).
  const rowFieldLabelsWithMeasure = valuesPlacement === "Rows" ? [...rowFieldLabels, "Σ Values"] : rowFieldLabels;
  const conditionalRanges = useMemo(() => {
    if (!result) return [];
    return valuesPlacement === "Rows"
      ? result.rowHeaders.map((_, r) => detailRangeForRow(result, r))
      : result.columnHeaders.map((_, c) => detailRangeForColumn(result, c));
  }, [result, valuesPlacement]);
  // Outline reuses the exact same visibility computation as Compact (see
  // computeCompactVisibility's own comment) — it just renders the result into separate <td> cells
  // per level instead of one merged, indented cell.
  const suppressedLevelVisibility = useMemo(
    () => (result && (layout === "Compact" || layout === "Outline") ? computeCompactVisibility(result.rowHeaders) : null),
    [result, layout],
  );

  if (!result) {
    return null;
  }

  const { rowHeaders, columnHeaders, cells } = result;
  const rowDepth = rowHeaders[0]?.labels.length ?? 0;
  const rowHeaderColumnCount = layout === "Compact" ? 1 : rowDepth;
  const columnDepth = columnHeaders[0]?.labels.length ?? 0;

  const toggleGroup = (key) => {
    setCollapsed((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  const groupForRow = new Map();
  groups.forEach((g) => {
    for (let r = g.start; r < g.end; r++) groupForRow.set(r, g);
  });

  const query = searchText.trim().toLowerCase();
  const matchesSearch = (row) => !query || row.isTotal || row.labels.some((l) => l.toLowerCase().includes(query));

  const isRowVisible = (row, rowIndex) => {
    const group = groupForRow.get(rowIndex);
    const isGroupSubtotalRow = group?.hasSubtotal && rowIndex === group.end - 1;
    const isCollapsedDetail = group?.hasSubtotal && collapsed.has(group.key) && !isGroupSubtotalRow;
    return !isCollapsedDetail && matchesSearch(row);
  };

  // Precomputed rather than mutated inline during the render map below, so numbering stays a
  // pure function of the current filters/collapse state instead of a running counter touched
  // mid-render.
  const visibleRowNumbers = new Map();
  {
    let counter = 0;
    rowHeaders.forEach((row, rowIndex) => {
      if (isRowVisible(row, rowIndex)) {
        counter += 1;
        visibleRowNumbers.set(rowIndex, counter);
      }
    });
  }

  return (
    // w-fit: the white card now hugs the table's actual (content-sized) width instead of
    // stretching to the full page — dropping the table's own w-full (previous fix) stopped
    // individual columns from being stretched, but left the surrounding card still full-width
    // with a large empty gap on the right, which is what "not sized right" meant here.
    <div className="card w-fit max-w-full overflow-hidden p-0">
      <div className="border-b border-slate-100 p-4">
        <div className="relative max-w-sm">
          <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted" />
          <input
            type="text"
            placeholder="Search rows…"
            value={searchText}
            onChange={(e) => setSearchText(e.target.value)}
            className="w-full rounded-full border border-slate-200 bg-white py-2 pl-9 pr-4 text-sm focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent"
          />
        </div>
      </div>

      <div className="overflow-x-auto">
        <table className="border-collapse text-sm">
          <thead>
            {columnFieldLabels.length > 0 && (
              <tr>
                <th colSpan={1 + rowHeaderColumnCount} />
                <th colSpan={columnHeaders.length} className="px-2 py-2 text-center">
                  <HeaderPill muted>{columnFieldLabels.join(" › ")}</HeaderPill>
                </th>
              </tr>
            )}
            {Array.from({ length: columnDepth }).map((_, level) => (
              <tr key={level}>
                {level === 0 && (
                  <>
                    <th rowSpan={columnDepth} className="px-2 py-2 text-center text-xs font-medium text-muted">
                      #
                    </th>
                    {layout === "Compact" ? (
                      <th rowSpan={columnDepth} className="px-2 py-2 text-center">
                        <HeaderPill muted>{rowFieldLabelsWithMeasure.join(" › ")}</HeaderPill>
                      </th>
                    ) : (
                      Array.from({ length: rowDepth }).map((_, rowLevel) => (
                        <th key={rowLevel} rowSpan={columnDepth} className="px-2 py-2 text-center">
                          <HeaderPill muted>{rowFieldLabelsWithMeasure[rowLevel] ?? ""}</HeaderPill>
                        </th>
                      ))
                    )}
                  </>
                )}
                {columnHeaders.map((col, colIndex) => (
                  <th key={colIndex} className={`px-2 py-2 text-center ${col.isTotal ? "bg-gold/10" : ""}`}>
                    <HeaderPill>{col.labels[level]}</HeaderPill>
                  </th>
                ))}
              </tr>
            ))}
          </thead>
          <tbody>
            {rowHeaders.map((row, rowIndex) => {
              if (!isRowVisible(row, rowIndex)) {
                return null;
              }

              const group = groupForRow.get(rowIndex);
              const isGroupSubtotalRow = group?.hasSubtotal && rowIndex === group.end - 1;
              const rowTint = row.isTotal ? "bg-gold/10" : visibleRowNumbers.get(rowIndex) % 2 === 0 ? "bg-slate-50/60" : "";

              const collapseToggle = isGroupSubtotalRow && (
                <button
                  type="button"
                  onClick={() => toggleGroup(group.key)}
                  className="text-muted hover:text-ink"
                  aria-label={collapsed.has(group.key) ? "Expand group" : "Collapse group"}
                >
                  {collapsed.has(group.key) ? <ChevronRight className="h-3.5 w-3.5" /> : <ChevronDown className="h-3.5 w-3.5" />}
                </button>
              );

              return (
                <tr key={rowIndex} className={`border-b border-slate-100 ${rowTint}`}>
                  <td className="whitespace-nowrap px-3 py-2.5 text-center text-xs text-muted">{visibleRowNumbers.get(rowIndex)}</td>
                  {layout === "Compact" ? (
                    <td
                      className={`whitespace-nowrap px-6 py-2.5 text-left ${row.isTotal ? "font-bold text-ink" : "font-medium text-ink"}`}
                    >
                      <span className="inline-flex items-center gap-1.5">
                        {collapseToggle}
                        <span>
                          {row.labels.map((label, level) =>
                            suppressedLevelVisibility[rowIndex][level] ? (
                              <div key={level} style={{ paddingLeft: level * 16 }}>
                                {label}
                              </div>
                            ) : null,
                          )}
                        </span>
                      </span>
                    </td>
                  ) : layout === "Outline" ? (
                    row.labels.map((label, level) => (
                      <td
                        key={level}
                        className={`whitespace-nowrap px-6 py-2.5 text-center ${row.isTotal ? "font-bold text-ink" : "font-medium text-ink"}`}
                      >
                        <span className="inline-flex items-center justify-center gap-1.5">
                          {level === 0 && collapseToggle}
                          {suppressedLevelVisibility[rowIndex][level] ? label : ""}
                        </span>
                      </td>
                    ))
                  ) : (
                    row.labels.map((label, level) => (
                      <td
                        key={level}
                        className={`whitespace-nowrap px-6 py-2.5 text-center ${row.isTotal ? "font-bold text-ink" : "font-medium text-ink"}`}
                      >
                        <span className="inline-flex items-center justify-center gap-1.5">
                          {level === 0 && collapseToggle}
                          {label}
                        </span>
                      </td>
                    ))
                  )}
                  {cells[rowIndex].map((cell, colIndex) => {
                    const isTotalCell = row.isTotal || columnHeaders[colIndex].isTotal;
                    // Which PivotValueField a cell belongs to is read from whichever axis actually
                    // carries the measure crossjoin — mirrors HtmlSnapshotRenderer.cs exactly.
                    const measureIndex = valuesPlacement === "Rows" ? rowIndex : colIndex;
                    const rangeIndex = valuesPlacement === "Rows" ? rowIndex : colIndex;
                    const valueField = valueFields?.length ? valueFields[measureIndex % valueFields.length] : null;
                    // Same rule as the saved HTML snapshot: Total/Subtotal cells never get a
                    // conditional-format background, since they're not "a value in the range" —
                    // they're the sum of it (see ConditionalFormatting.cs for the full reasoning).
                    const conditionalStyle = !isTotalCell && valueField
                      ? conditionalStyleFor(cell, valueField.conditionalFormat, conditionalRanges[rangeIndex]?.min, conditionalRanges[rangeIndex]?.max)
                      : null;
                    const iconColor = !isTotalCell && valueField
                      ? iconColorFor(cell, valueField.conditionalFormat, conditionalRanges[rangeIndex]?.min, conditionalRanges[rangeIndex]?.max)
                      : null;

                    return (
                      <td
                        key={colIndex}
                        className={`whitespace-nowrap px-6 py-2.5 text-center tabular-nums ${
                          isTotalCell ? "font-bold text-ink" : "text-slate-700"
                        }`}
                        style={conditionalStyle ?? undefined}
                      >
                        {iconColor && <span style={{ color: iconColor }}>●</span>} {formatNumber(cell, formatForColumn(measureIndex, valueFields))}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <p className="border-t border-slate-100 px-4 py-3 text-xs text-muted">
        {[...visibleRowNumbers.values()].length.toLocaleString()} of {rowHeaders.length.toLocaleString()} rows
      </p>
    </div>
  );
}
