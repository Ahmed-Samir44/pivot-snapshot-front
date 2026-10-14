// Shared native-HTML5-DnD payload (no library) for moving a field between any of the pivot
// builder's zones — Filters, Rows, Columns, Values, and the Available Fields source list.
// dataTransfer carries a plain JSON string across the browser's DOM event system, which is what
// lets completely separate React component instances (each zone owns its own state slice)
// coordinate a move without any shared component tree above FieldPicker.
export const FIELD_DRAG_MIME = "application/x-pivot-field";

export function setFieldDragPayload(event, field, sourceZone) {
  event.dataTransfer.setData(FIELD_DRAG_MIME, JSON.stringify({ field, source: sourceZone }));
  event.dataTransfer.effectAllowed = "move";
}

export function readFieldDragPayload(event) {
  const raw = event.dataTransfer.getData(FIELD_DRAG_MIME);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

// A CalculatedItem's Field must still be a Rows or Columns field — the backend rejects one that
// isn't (see MdxPivotQueryBuilder.ValidateCalculatedItems) — so it needs to be dropped the moment
// its field leaves both axes, same idea as dateGroupings/numericGroupings being dropped when their
// field leaves Rows.
export function clearOrphanedCalculatedItems(calculatedItems, rows, columns) {
  return calculatedItems.filter((item) => rows.includes(item.field) || columns.includes(item.field));
}

// "[Dim].[Hierarchy].[Level]" → "[Dim].[Hierarchy]" (unchanged if it's already 2 segments).
function hierarchyOf(field) {
  return field.match(/^(\[[^\]]+\]\.\[[^\]]+\])/)?.[1] ?? field;
}

// Drilldown v3 (2026-10-14): the drilldown hierarchy is a SLOT among the Rows fields (its own
// topLevelField sits in Rows). It survives any Rows change as long as that slot is still there and
// no other Rows field comes from the same hierarchy (SSAS can't put one hierarchy on an axis twice);
// otherwise drilldown mode ends — the slot is dropped too, so no stale entry is left in Rows.
// Returns the reconciled { rows, rowsHierarchyDrilldown } pair to spread into the next request.
export function reconcileDrilldown(drilldown, rows) {
  if (!drilldown) return { rows, rowsHierarchyDrilldown: null };
  const top = drilldown.topLevelField;
  const clash = rows.some((f) => f !== top && hierarchyOf(f) === hierarchyOf(top));
  if (rows.includes(top) && !clash) return { rows, rowsHierarchyDrilldown: drilldown };
  return { rows: rows.filter((f) => f !== top), rowsHierarchyDrilldown: null };
}

// Several request pieces reference a Value by its bare FIELD NAME (Filter.byMeasureField for
// TopN/BottomN/GreaterThan/LessThan/Between, Sort.byMeasureField, RowLevelSorts/ColumnLevelSorts'
// own byMeasureField) — removing that Value from Values used to leave these references dangling,
// which the backend rejects on the very next query with no indication of which setting is stale
// (caught 2026-10-06, full codebase review: e.g. sort Rows by value of "Revenue", then drag
// "Revenue" out of Values — the axis Sort still says byMeasureField: "Revenue"). Mirrors
// clearOrphanedCalculatedItems's own "drop/reset whatever referenced the thing that's gone" idea.
export function clearDanglingMeasureReferences(request, remainingValueFields) {
  const stillExists = (field) => remainingValueFields.includes(field);
  return {
    ...request,
    // A filter's OWN field stays — only its value-dependent mode gets reset to a safe default,
    // same shape FilterRow's own handleModeChange uses for any other mode switch.
    filters: request.filters.map((f) =>
      f.byMeasureField && !stillExists(f.byMeasureField)
        ? { field: f.field, mode: "Members", includedMembers: [], labelText: "", n: 10, byMeasureField: null, value: null, valueTo: null }
        : f,
    ),
    sort: request.sort?.byMeasureField && !stillExists(request.sort.byMeasureField) ? null : request.sort,
    rowLevelSorts: request.rowLevelSorts.filter((s) => !s.byMeasureField || stillExists(s.byMeasureField)),
    columnLevelSorts: request.columnLevelSorts.filter((s) => !s.byMeasureField || stillExists(s.byMeasureField)),
  };
}

// Same idea as clearDanglingMeasureReferences, for a Value's own baseField (Show Values As
// PercentOf/DifferenceFrom/PercentDifferenceFrom) — baseField/baseItem target a ROWS field, not a
// measure, so this triggers when a ROWS field (not a Value) is removed instead.
export function clearDanglingBaseFieldReferences(values, remainingRowFields) {
  return values.map((v) =>
    v.baseField && !remainingRowFields.includes(v.baseField) ? { ...v, baseField: null, baseItem: "" } : v,
  );
}

// Excel's own "drag a field out of the Field List entirely and it's removed" — not just dragging
// it back onto the Available Fields panel specifically (FieldPicker's own moveField "available"
// case), but dropping it ANYWHERE on the page outside every recognized zone. Lives here (not in
// FieldPicker.jsx, which only exports components) so PivotBuilder.jsx's own page-level catch-all
// drop handler can reuse the exact same removal — every zone's own onDrop already stops
// propagation before reaching that handler (see each one's own WHY comment), so it only ever fires
// for a drop that landed nowhere FieldPicker itself recognizes.
export function removeFieldEverywhere(request, field) {
  const rows = request.rows.filter((f) => f !== field);
  const columns = request.columns.filter((f) => f !== field);
  // Same reasoning as moveField's own WHY comment in FieldPicker.jsx — a removed field shifts every
  // OTHER level-sort entry's own positional `level` index, so a stale rowLevelSorts/columnLevelSorts
  // left over from before this removal would point at the wrong field (or an out-of-range one).
  return {
    ...request,
    rows,
    columns,
    values: request.values.filter((v) => v.field !== field),
    filters: request.filters.filter((f) => f.field !== field),
    dateGroupings: request.dateGroupings.filter((g) => g.field !== field),
    numericGroupings: request.numericGroupings.filter((g) => g.field !== field),
    calculatedItems: clearOrphanedCalculatedItems(request.calculatedItems, rows, columns),
    rowLevelSorts: request.rows.includes(field) ? [] : request.rowLevelSorts,
    columnLevelSorts: request.columns.includes(field) ? [] : request.columnLevelSorts,
  };
}
