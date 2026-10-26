import { isFilterComplete } from "./filterModes";
import { isCompleteCalculatedField } from "./calculatedFields";

// Excel's slicer greys out the items that have no data under the OTHER filters on the PivotTable
// and lists them last ("Visually indicate items with no data" / "Show items with no data last", on
// by default — 2026-10-26). This builds the small query that answers "which items of this field
// have data": just the field on Rows with the first value field, every filter except the slicer's
// own, and no totals, sorting or groupings. Null when the PivotTable has no value field yet.
export function buildSlicerProbeRequest(request, field) {
  const value = request?.values?.find((v) => v.field);
  if (!value) return null;
  const { format: _format, conditionalFormat: _conditionalFormat, ...measure } = value;
  return {
    rows: [field],
    columns: [],
    values: [measure],
    filters: request.filters.filter((f) => f.field && f.field !== field && isFilterComplete(f)),
    sort: null,
    rowLevelSorts: [],
    columnLevelSorts: [],
    grandTotals: "Off",
    subtotals: "Off",
    includeFilteredItemsInTotals: false,
    allowMultipleFiltersPerField: Boolean(request.allowMultipleFiltersPerField),
    fieldSubtotals: [],
    calculatedFields: (request.calculatedFields ?? []).filter(isCompleteCalculatedField),
    calculatedItems: [],
    dateGroupings: [],
    numericGroupings: [],
    itemGroups: [],
    showItemsWithNoData: [],
    fieldsShowingItemsWithNoData: [],
    rowsHierarchyDrilldown: null,
    columnsHierarchyDrilldown: null,
  };
}

// The items that came back with a value — what the slicer shows as "has data".
export function itemsWithData(result) {
  const items = new Set();
  (result?.rowHeaders ?? []).forEach((header, r) => {
    if (header.isTotal) return;
    const hasValue = (result.cells?.[r] ?? []).some((v) => typeof v === "number" && Number.isFinite(v));
    if (hasValue) items.add(header.labels[0] ?? "");
  });
  return items;
}
