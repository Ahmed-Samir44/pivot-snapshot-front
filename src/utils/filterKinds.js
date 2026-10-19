import { isDateMode, isLabelMode } from "./filterModes";

// Excel's "Allow multiple filters per field" (2026-10-19): a field can hold one filter of each KIND
// at once — 0 = a manual selection (Members/Exclude), 1 = a label filter, 2 = a value filter —
// applied in that order. Mirrors MdxPivotQueryBuilder.FilterOrder.
export function filterKind(filter) {
  const mode = filter.mode ?? "Members";
  if (mode === "Members" || mode === "Exclude") return 0;
  // A date filter is the same kind as a label filter (Excel shows one in place of the other).
  if (isLabelMode(mode) || isDateMode(mode)) return 1;
  return 2;
}

export const FILTER_KIND_LABELS = ["Manual selection", "Label filter", "Value filter"];

// Puts `filter` on its field: with multiple filters allowed it replaces only the field's filter of
// the same kind, otherwise every filter on the field (one per field — Excel's default).
export function setFieldFilter(filters, filter, allowMultiple) {
  const replaced = (f) => f.field === filter.field && (!allowMultiple || filterKind(f) === filterKind(filter));
  const index = filters.findIndex(replaced);
  const kept = filters.filter((f) => !replaced(f));
  // Keep its place in the list when it replaced one (the Report Filter area's order), else append.
  return index < 0 ? [...kept, filter] : [...kept.slice(0, index), filter, ...kept.slice(index)];
}

// One filter per field — what turning "Allow multiple filters per field" off leaves: the field's
// first filter in Excel's order (a manual selection before a label filter before a value filter).
export function onePerField(filters) {
  const best = new Map();
  filters.forEach((f) => {
    const current = best.get(f.field);
    if (!current || filterKind(f) < filterKind(current)) best.set(f.field, f);
  });
  return filters.filter((f) => best.get(f.field) === f);
}
