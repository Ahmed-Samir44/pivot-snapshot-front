// Ported from the Segmentation project's src/utils/constants.js — generic MDX bracket-path
// parsing/formatting, not project-specific logic, so reusing it keeps both tools' field pickers
// behaving identically for the same kind of input.

/** Split an MDX bracket path into each segment, e.g. "[A].[B].[C]" -> ["A", "B", "C"]. */
export function parseCubeBracketSegments(path) {
  const matches = String(path).match(/\[([^\]]*)\]/g);
  return matches ? matches.map((s) => s.slice(1, -1)) : [];
}

/** Friendly label for a single bracket segment (used for picker dropdown steps). */
export function formatCubeBracketLabel(segment) {
  if (segment == null || segment === "") return "";
  return String(segment)
    .trim()
    .replace(/^Dim\s+/i, "")
    .replace(/_/g, " ");
}

/** Friendly label for a full bracket path, deduping a repeated trailing segment. */
export function formatDimensionName(dimension) {
  const parts = String(dimension)
    .replace(/\[Dim\s+/g, "[")
    .replace(/\[|\]/g, "")
    .replace(/\s*-\s*/g, " - ")
    .split(".")
    .filter((part, index, arr) => part !== arr[index - 1]);
  return parts.join(".");
}
