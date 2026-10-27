// An error cell is kept INSIDE its own cell so it travels through every sort/total/layout step (see
// the backend's CubeQueryService.ParseCellValue). Two kinds, as in Excel:
//   #DIV/0! — a divide-by-zero in a calculated field/item or a % — is NaN;
//   #N/A    — a Show Values As base item the row's group doesn't have (2026-10-27, see the
//             backend's ShowValuesAsProcessor.NotAvailable) — is +Infinity.
// JSON has neither, so the backend writes them as the strings "NaN" and "Infinity"; these convert at
// the API boundary — JSON.stringify would otherwise turn them into null (a blank).

export const NOT_AVAILABLE = Infinity;

export function isErrorValue(value) {
  return typeof value === "number" && (Number.isNaN(value) || value === NOT_AVAILABLE);
}

// Excel's text for an error cell.
export function errorToken(value) {
  return value === NOT_AVAILABLE ? "#N/A" : "#DIV/0!";
}

const reviveToken = (v) => (v === "Infinity" ? NOT_AVAILABLE : v === "NaN" || v === "-Infinity" ? NaN : v);

// "NaN"/"Infinity" strings → their numbers (returns a new result object).
export function reviveResultCells(result) {
  if (!result?.cells) return result;
  return { ...result, cells: result.cells.map((row) => row.map(reviveToken)) };
}

// Error numbers → "NaN"/"Infinity" strings, for sending a result back to the backend (saving a snapshot).
export function serializeResultCells(result) {
  if (!result?.cells) return result;
  return { ...result, cells: result.cells.map((row) => row.map((v) => (v === NOT_AVAILABLE ? "Infinity" : isErrorValue(v) ? "NaN" : v))) };
}
