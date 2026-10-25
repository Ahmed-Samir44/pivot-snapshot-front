import { useMemo, useState } from "react";
import { isErrorValue } from "../../utils/errorCells";

// Side-by-side comparison of 2+ saved versions (2026-10-20 — the backend's /compare endpoint had no
// screen). One row per cell of the pivot (row labels × column labels, matched by LABEL across
// versions — SnapshotComparer), a column per version oldest → newest, then the change from the
// first to the last version as a number and a percentage. "Changed only" hides cells whose value is
// the same in every version.

const numberFormat = new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 });

function formatValue(value) {
  if (value == null) return "";
  if (isErrorValue(value)) return "#DIV/0!";
  return numberFormat.format(value);
}

// A Grand Total row/column is "Total" at every level — shown as one "Grand Total".
const joinLabels = (labels) => {
  const shown = labels.filter((l) => l !== "");
  return shown.length > 0 && shown.every((l) => l === "Total") ? "Grand Total" : shown.join(" › ");
};

// The cube's sums come back with float noise (1,924,672,814.28 vs …814.2800001) — equal within a
// relative 1e-9 counts as unchanged, or every cell would show as "+0" changed.
const sameValue = (a, b) => {
  if (a == null || b == null) return a == null && b == null;
  if (isErrorValue(a) || isErrorValue(b)) return isErrorValue(a) && isErrorValue(b);
  return Math.abs(a - b) <= 1e-9 * Math.max(Math.abs(a), Math.abs(b), 1);
};

export default function VersionCompare({ comparison, versionLabels }) {
  const [changedOnly, setChangedOnly] = useState(true);
  const [search, setSearch] = useState("");

  // Only one distinct column (a single measure, no Columns fields): the Column header adds nothing.
  const showColumn = useMemo(() => new Set(comparison.cells.map((c) => joinLabels(c.columnLabels))).size > 1, [comparison]);

  const rows = useMemo(() => {
    const term = search.trim().toLowerCase();
    return comparison.cells
      .map((cell) => {
        const first = cell.values[0];
        const last = cell.values[cell.values.length - 1];
        const usable = (v) => v != null && !isErrorValue(v);
        const changed = cell.values.some((v) => !sameValue(v, cell.values[0]));
        const change = usable(first) && usable(last) ? (sameValue(first, last) ? 0 : last - first) : null;
        const percent = change != null && first !== 0 ? change / Math.abs(first) : null;
        // An item in only some of the versions (a department that appeared or went away).
        const presence = first == null && last != null ? "New" : first != null && last == null ? "Removed" : null;
        return { cell, change, percent, changed, presence, rowLabel: joinLabels(cell.rowLabels), columnLabel: joinLabels(cell.columnLabels) };
      })
      .filter((r) => !changedOnly || r.changed)
      .filter((r) => !term || r.rowLabel.toLowerCase().includes(term) || r.columnLabel.toLowerCase().includes(term));
  }, [comparison, changedOnly, search]);

  const changedCount = useMemo(() => comparison.cells.filter((c) => c.values.some((v) => !sameValue(v, c.values[0]))).length, [comparison]);

  const changeClass = (v) => (v == null || v === 0 ? "text-muted" : v > 0 ? "text-green-700" : "text-red-700");

  // A pivot with no Values (row labels only) has no cells to line up.
  if (comparison.cells.length === 0) {
    return <p className="text-sm text-muted">These versions have no Values — only row labels — so there are no numbers to compare.</p>;
  }

  return (
    <div className="space-y-3">
      <div className="flex flex-wrap items-center gap-3 text-sm">
        <input
          type="search"
          className="input-field w-56 py-1 text-sm"
          placeholder="Search labels…"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
        />
        <label className="flex cursor-pointer items-center gap-2">
          <input type="checkbox" className="h-4 w-4 accent-primary" checked={changedOnly} onChange={(e) => setChangedOnly(e.target.checked)} />
          Changed only
        </label>
        <span className="text-xs text-muted">
          {changedCount} of {comparison.cells.length} cells changed
        </span>
      </div>

      <div className="custom-scrollbar max-h-[65vh] overflow-auto rounded-lg border border-slate-200">
        <table className="min-w-full border-collapse text-sm">
          <thead className="sticky top-0 bg-slate-50">
            <tr>
              <th className="border-b border-slate-200 px-3 py-2 text-left font-semibold text-ink">Row</th>
              {showColumn && <th className="border-b border-slate-200 px-3 py-2 text-left font-semibold text-ink">Column</th>}
              {versionLabels.map((label) => (
                <th key={label} className="whitespace-nowrap border-b border-slate-200 px-3 py-2 text-right font-semibold text-ink">
                  {label}
                </th>
              ))}
              <th className="whitespace-nowrap border-b border-slate-200 px-3 py-2 text-right font-semibold text-ink">Change</th>
              <th className="whitespace-nowrap border-b border-slate-200 px-3 py-2 text-right font-semibold text-ink">Change %</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(({ cell, change, percent, presence, rowLabel, columnLabel }, i) => (
              <tr key={i} className="odd:bg-white even:bg-slate-50/60">
                <td className="px-3 py-1.5 text-ink">{rowLabel || "Grand Total"}</td>
                {showColumn && <td className="px-3 py-1.5 text-muted">{columnLabel}</td>}
                {cell.values.map((v, j) => (
                  <td key={j} className="whitespace-nowrap px-3 py-1.5 text-right tabular-nums">
                    {v == null ? <span className="text-muted">—</span> : formatValue(v)}
                  </td>
                ))}
                <td className={`whitespace-nowrap px-3 py-1.5 text-right font-medium tabular-nums ${presence ? (presence === "New" ? "text-green-700" : "text-red-700") : changeClass(change)}`}>
                  {presence ?? (change == null ? "" : `${change > 0 ? "+" : ""}${formatValue(change)}`)}
                </td>
                <td className={`whitespace-nowrap px-3 py-1.5 text-right tabular-nums ${changeClass(percent)}`}>
                  {percent == null ? "" : `${percent > 0 ? "+" : ""}${(percent * 100).toFixed(1)}%`}
                </td>
              </tr>
            ))}
            {rows.length === 0 && (
              <tr>
                <td colSpan={versionLabels.length + (showColumn ? 4 : 3)} className="px-3 py-6 text-center text-muted">
                  {changedOnly ? "No differences between these versions." : "Nothing matches."}
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
