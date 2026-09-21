import { useMemo, useState } from "react";
import { ChevronDown, ChevronRight } from "lucide-react";

// Groups consecutive rows sharing the same first-level label into collapsible sections. A group
// is only collapsible when it ends in a Subtotal row (IsTotal, added by the backend's
// SubtotalsProcessor when Rows has 2+ fields and Show Subtotals is on) — collapsing without a
// subtotal to fall back to would just hide data with nothing left summarizing it, so no toggle is
// shown in that case (matches how Excel only offers +/- once subtotals exist).
function computeGroups(rowHeaders) {
  const groups = [];
  let i = 0;
  while (i < rowHeaders.length) {
    if (rowHeaders[i].isTotal) {
      // The Grand Total row (or a stray total with no preceding group) — not collapsible.
      groups.push({ start: i, end: i + 1, key: null, hasSubtotal: false });
      i += 1;
      continue;
    }
    const start = i;
    const key = rowHeaders[i].labels[0];
    while (i < rowHeaders.length && !rowHeaders[i].isTotal && rowHeaders[i].labels[0] === key) {
      i += 1;
    }
    const hasSubtotal = i < rowHeaders.length && rowHeaders[i].isTotal;
    const end = hasSubtotal ? i + 1 : i;
    if (hasSubtotal) i += 1;
    groups.push({ start, end, key, hasSubtotal });
  }
  return groups;
}

export default function PivotGrid({ result }) {
  const [collapsed, setCollapsed] = useState(() => new Set());

  const groups = useMemo(() => (result ? computeGroups(result.rowHeaders) : []), [result]);

  if (!result) {
    return null;
  }

  const { rowHeaders, columnHeaders, cells } = result;
  const rowDepth = rowHeaders[0]?.labels.length ?? 0;
  const columnDepth = columnHeaders[0]?.labels.length ?? 0;

  const totalCellClass = "bg-gold/20 font-bold";
  const numberColClass = "border border-slate-200 bg-slate-100 px-2 py-2 text-center text-xs text-muted";

  const toggleGroup = (key) => {
    setCollapsed((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  // A detail row is hidden when it belongs to a collapsed group and isn't that group's own
  // Subtotal row (the Subtotal row always stays visible — it's the collapsed summary).
  const groupForRow = new Map();
  groups.forEach((g) => {
    for (let r = g.start; r < g.end; r++) groupForRow.set(r, g);
  });

  // Precomputed rather than mutated inline during the render map below, so numbering stays a
  // pure function of `rowHeaders`/`collapsed` instead of a running counter touched mid-render.
  const visibleRowNumbers = new Map();
  {
    let counter = 0;
    rowHeaders.forEach((row, rowIndex) => {
      const group = groupForRow.get(rowIndex);
      const isGroupSubtotalRow = group?.hasSubtotal && rowIndex === group.end - 1;
      const isCollapsedDetail = group?.hasSubtotal && collapsed.has(group.key) && !isGroupSubtotalRow;
      if (!isCollapsedDetail) {
        counter += 1;
        visibleRowNumbers.set(rowIndex, counter);
      }
    });
  }

  return (
    <div className="card overflow-x-auto">
      <table className="w-full border-collapse text-sm">
        <thead>
          {Array.from({ length: columnDepth }).map((_, level) => (
            <tr key={level}>
              {level === 0 && (
                <>
                  <th rowSpan={columnDepth} className={numberColClass}>
                    #
                  </th>
                  <th colSpan={rowDepth} rowSpan={columnDepth} className="border border-slate-200" />
                </>
              )}
              {columnHeaders.map((col, colIndex) => (
                <th
                  key={colIndex}
                  className={`border border-slate-200 bg-gold/10 px-3 py-2 text-ink ${col.isTotal ? totalCellClass : ""}`}
                >
                  {col.labels[level]}
                </th>
              ))}
            </tr>
          ))}
        </thead>
        <tbody>
          {rowHeaders.map((row, rowIndex) => {
            const group = groupForRow.get(rowIndex);
            const isGroupSubtotalRow = group?.hasSubtotal && rowIndex === group.end - 1;
            const isCollapsedDetail = group?.hasSubtotal && collapsed.has(group.key) && !isGroupSubtotalRow;
            if (isCollapsedDetail) {
              return null;
            }

            return (
              <tr key={rowIndex}>
                <td className={numberColClass}>{visibleRowNumbers.get(rowIndex)}</td>
                {row.labels.map((label, level) => (
                  <th
                    key={level}
                    scope="row"
                    className={`border border-slate-200 bg-gold/10 px-3 py-2 text-left text-ink ${row.isTotal ? totalCellClass : ""}`}
                  >
                    <span className="inline-flex items-center gap-1">
                      {level === 0 && isGroupSubtotalRow && (
                        <button
                          type="button"
                          onClick={() => toggleGroup(group.key)}
                          className="text-muted hover:text-ink"
                          aria-label={collapsed.has(group.key) ? "Expand group" : "Collapse group"}
                        >
                          {collapsed.has(group.key) ? <ChevronRight className="h-3.5 w-3.5" /> : <ChevronDown className="h-3.5 w-3.5" />}
                        </button>
                      )}
                      {label}
                    </span>
                  </th>
                ))}
                {cells[rowIndex].map((cell, colIndex) => (
                  <td
                    key={colIndex}
                    className={`border border-slate-200 px-3 py-2 text-right ${
                      row.isTotal || columnHeaders[colIndex].isTotal ? totalCellClass : ""
                    }`}
                  >
                    {cell ?? ""}
                  </td>
                ))}
              </tr>
            );
          })}
        </tbody>
      </table>
      <p className="mt-2 text-xs text-muted">{rowHeaders.length.toLocaleString()} rows</p>
    </div>
  );
}
