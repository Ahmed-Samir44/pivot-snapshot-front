// An error cell (Excel's #DIV/0! — a divide-by-zero in a calculated field/item) is NaN, kept INSIDE
// its own cell so it travels through every sort/total/layout step (see the backend's
// CubeQueryService.ParseCellValue). JSON has no NaN, so the backend writes it as the string "NaN";
// these convert at the API boundary — JSON.stringify would otherwise turn NaN into null (a blank).

const isErrorToken = (v) => v === "NaN" || v === "Infinity" || v === "-Infinity";

export function isErrorValue(value) {
  return typeof value === "number" && Number.isNaN(value);
}

// "NaN" strings → NaN numbers (returns a new result object).
export function reviveResultCells(result) {
  if (!result?.cells) return result;
  return { ...result, cells: result.cells.map((row) => row.map((v) => (isErrorToken(v) ? NaN : v))) };
}

// NaN numbers → "NaN" strings, for sending a result back to the backend (saving a snapshot).
export function serializeResultCells(result) {
  if (!result?.cells) return result;
  return { ...result, cells: result.cells.map((row) => row.map((v) => (isErrorValue(v) ? "NaN" : v))) };
}
