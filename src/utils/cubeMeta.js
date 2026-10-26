// Ported from the Segmentation project's src/utils/constants.js — generic MDX bracket-path
// parsing/formatting, not project-specific logic, so reusing it keeps both tools' field pickers
// behaving identically for the same kind of input.
import { MEASURES_PLACEHOLDER } from "./measuresPlacement";

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
  // MEASURES_PLACEHOLDER is a reserved sentinel, not a real cube field — Excel's own literal
  // label for the "Σ Values" pseudo-field, wherever it's placed (mirrors the backend's
  // FieldLabelFormatter.Format exactly).
  if (dimension === MEASURES_PLACEHOLDER) {
    return "Σ Values";
  }

  const parts = String(dimension)
    .replace(/\[Dim\s+/g, "[")
    .replace(/\[|\]/g, "")
    .replace(/\s*-\s*/g, " - ")
    .split(".")
    .filter((part, index, arr) => part !== arr[index - 1]);
  return parts.join(".");
}

// A field's own caption, the way Excel labels an OLAP field in the PivotTable — "BU Short Name",
// not "Business Unit.BU Short Name" (2026-10-26). formatDimensionName keeps the full
// "Table.Field" name for tooltips and for telling same-named fields apart (see fieldCaptions).
// Accepts a raw MDX field id or an already formatted "Table.Field" display name.
export function fieldCaption(dimension) {
  if (dimension === MEASURES_PLACEHOLDER) return "Σ Values";
  const text = String(dimension ?? "");
  const segments = parseCubeBracketSegments(text);
  if (segments.length > 0) return segments[segments.length - 1];
  return text.includes(".") ? text.slice(text.lastIndexOf(".") + 1) : text;
}

// Captions for the fields of one PivotTable: each field's short caption, except where two
// different fields would read the same ("Date" from two date tables) — those keep their full
// "Table.Field" name so the headers stay unambiguous. fullNameOf(id) gives that full name.
export function fieldCaptions(ids, fullNameOf = formatDimensionName) {
  const unique = [...new Set(ids.filter((id) => id && id !== MEASURES_PLACEHOLDER))];
  const counts = new Map();
  for (const id of unique) counts.set(fieldCaption(fullNameOf(id)), (counts.get(fieldCaption(fullNameOf(id))) ?? 0) + 1);
  return (id) => {
    if (id === MEASURES_PLACEHOLDER) return "Σ Values";
    const caption = fieldCaption(fullNameOf(id));
    return (counts.get(caption) ?? 0) > 1 ? fullNameOf(id) : caption;
  };
}
