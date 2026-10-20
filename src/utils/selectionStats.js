import { isErrorValue } from "./errorCells";

// Excel's status-bar summary of the selected cells (2026-10-20): Count = every non-empty cell (an
// error cell included), Sum/Average over the numeric cells only. null when nothing is selected.
// average/sum are null when the range holds no numbers (Excel then shows just the Count).
export function selectionStats(values) {
  if (!values || values.length === 0) return null;
  let count = 0;
  let numericCount = 0;
  let sum = 0;
  for (const value of values) {
    if (value === null || value === undefined || value === "") continue;
    count += 1;
    if (typeof value === "number" && !isErrorValue(value) && Number.isFinite(value)) {
      numericCount += 1;
      sum += value;
    }
  }
  return {
    cellCount: values.length,
    count,
    numericCount,
    sum: numericCount > 0 ? sum : null,
    average: numericCount > 0 ? sum / numericCount : null,
  };
}
