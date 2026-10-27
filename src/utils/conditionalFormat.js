// Mirrors the backend's ConditionalFormatting.cs exactly, so the live grid shows the same
// colors/bars the saved HTML snapshot will end up with. Display-only, computed from whatever
// detail cells are currently in the result — never sent to or computed by the cube.
//
// DEFAULT_MIN/MID/MAX/BAR are Excel's own stock Red-Yellow-Green scale and data-bar blue — the
// fallback whenever a ConditionalFormat doesn't specify its own colors (minColor/midColor/maxColor/
// barColor, added so the user can customize them, same as real Excel's own Color Scale/Data Bar
// dialogs let you pick your own colors instead of being stuck with the one preset).
export const DEFAULT_MIN_COLOR = "#F8696B";
export const DEFAULT_MID_COLOR = "#FFEB84";
export const DEFAULT_MAX_COLOR = "#63BE7B";
export const DEFAULT_BAR_COLOR = "#638EC6";

function clamp01(value) {
  return Math.min(Math.max(value, 0), 1);
}

function hexToRgb(hex) {
  return [parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16)];
}

function toHex(n) {
  return n.toString(16).padStart(2, "0").toUpperCase();
}

function colorScaleHex(t, format) {
  const min = hexToRgb(format?.minColor || DEFAULT_MIN_COLOR);
  const mid = hexToRgb(format?.midColor || DEFAULT_MID_COLOR);
  const max = hexToRgb(format?.maxColor || DEFAULT_MAX_COLOR);
  const [from, to, localT] = t < 0.5 ? [min, mid, t * 2] : [mid, max, (t - 0.5) * 2];
  const r = Math.round(from[0] + (to[0] - from[0]) * localT);
  const g = Math.round(from[1] + (to[1] - from[1]) * localT);
  const b = Math.round(from[2] + (to[2] - from[2]) * localT);
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

// Returns a React inline-style object (or null) — the caller spreads it into the cell's `style`.
export function conditionalStyleFor(value, format, min, max) {
  // An error cell (NaN) gets no formatting, like Excel.
  if (!format || format.type === "None" || value === null || value === undefined || !Number.isFinite(value)) {
    return null;
  }

  // max <= min means every visible cell in range has the SAME value (the common case: a filter
  // leaves one row, or a collapsed group shows just its own subtotal) — (value-min)/(max-min) is
  // 0/0 there, not a real position on the scale. Real Excel's Color Scale/Data Bar/Icon Set don't
  // omit formatting for this case, they still render something consistent — midpoint (t = 0.5) is
  // the most visually-neutral choice (neither extreme), confirmed live 2026-10-06 ("اي حاجة مش
  // فارق الالوان... شوف الصح ايه") rather than the blank cell this used to fall back to.
  const t = max <= min ? 0.5 : clamp01((value - min) / (max - min));

  if (format.type === "ColorScale") {
    return { background: colorScaleHex(t, format) };
  }
  if (format.type === "DataBar") {
    const pct = Math.round(t * 100);
    const barColor = format.barColor || DEFAULT_BAR_COLOR;
    return { background: `linear-gradient(to right, ${barColor} ${pct}%, transparent ${pct}%)` };
  }
  return null;
}

// Excel's default "3 Traffic Lights" tertile split — same colors as the color scale (customizable
// the same way), mirrors the backend's ConditionalFormatting.IconColorFor exactly. Returns just the
// hex color; the caller renders the actual dot (see PivotGrid.jsx).
export function iconColorFor(value, format, min, max) {
  if (!format || format.type !== "IconSet" || value === null || value === undefined || !Number.isFinite(value)) {
    return null;
  }

  // Same max <= min / midpoint reasoning as conditionalStyleFor above.
  const t = max <= min ? 0.5 : clamp01((value - min) / (max - min));
  if (t < 1 / 3) return format.minColor || DEFAULT_MIN_COLOR;
  if (t < 2 / 3) return format.midColor || DEFAULT_MID_COLOR;
  return format.maxColor || DEFAULT_MAX_COLOR;
}

// Same "detail rows only" exclusion as the backend's DetailRangeFor — a Grand Total is always the
// largest number in its column, so including it would stretch every real value toward the "low"
// end of the scale.
export function detailRangeForColumn(result, columnIndex) {
  const values = result.rowHeaders
    .map((row, r) => (row.isTotal ? null : result.cells[r][columnIndex]))
    .filter((v) => v !== null && v !== undefined && Number.isFinite(v)); // error cells (#DIV/0!, #N/A) don't count

  if (values.length === 0) {
    return { min: 0, max: 0 };
  }
  return { min: Math.min(...values), max: Math.max(...values) };
}

// Mirror of detailRangeForColumn, across the OTHER axis — used whenever Σ Values sits in Rows
// (see PivotGrid.jsx), where a single ROW (not column) holds one measure's whole detail range.
export function detailRangeForRow(result, rowIndex) {
  const values = result.columnHeaders
    .map((col, c) => (col.isTotal ? null : result.cells[rowIndex][c]))
    .filter((v) => v !== null && v !== undefined && Number.isFinite(v)); // error cells (#DIV/0!, #N/A) don't count

  if (values.length === 0) {
    return { min: 0, max: 0 };
  }
  return { min: Math.min(...values), max: Math.max(...values) };
}
