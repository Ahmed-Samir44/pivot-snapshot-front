// Mirrors the backend's ConditionalFormatting.cs exactly, so the live grid shows the same
// colors/bars the saved HTML snapshot will end up with. Display-only, computed from whatever
// detail cells are currently in the result — never sent to or computed by the cube.
const RED = [0xf8, 0x69, 0x6b];
const YELLOW = [0xff, 0xeb, 0x84];
const GREEN = [0x63, 0xbe, 0x7b];
const DATA_BAR_COLOR = "#638EC6";

function clamp01(value) {
  return Math.min(Math.max(value, 0), 1);
}

function toHex(n) {
  return n.toString(16).padStart(2, "0").toUpperCase();
}

function colorScaleHex(t) {
  const [from, to, localT] = t < 0.5 ? [RED, YELLOW, t * 2] : [YELLOW, GREEN, (t - 0.5) * 2];
  const r = Math.round(from[0] + (to[0] - from[0]) * localT);
  const g = Math.round(from[1] + (to[1] - from[1]) * localT);
  const b = Math.round(from[2] + (to[2] - from[2]) * localT);
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

// Returns a React inline-style object (or null) — the caller spreads it into the cell's `style`.
export function conditionalStyleFor(value, format, min, max) {
  if (!format || format.type === "None" || value === null || value === undefined || max <= min) {
    return null;
  }

  const t = clamp01((value - min) / (max - min));

  if (format.type === "ColorScale") {
    return { background: colorScaleHex(t) };
  }
  if (format.type === "DataBar") {
    const pct = Math.round(t * 100);
    return { background: `linear-gradient(to right, ${DATA_BAR_COLOR} ${pct}%, transparent ${pct}%)` };
  }
  return null;
}

// Same "detail rows only" exclusion as the backend's DetailRangeFor — a Grand Total is always the
// largest number in its column, so including it would stretch every real value toward the "low"
// end of the scale.
export function detailRangeForColumn(result, columnIndex) {
  const values = result.rowHeaders
    .map((row, r) => (row.isTotal ? null : result.cells[r][columnIndex]))
    .filter((v) => v !== null && v !== undefined);

  if (values.length === 0) {
    return { min: 0, max: 0 };
  }
  return { min: Math.min(...values), max: Math.max(...values) };
}
