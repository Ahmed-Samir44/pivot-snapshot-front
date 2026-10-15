// Excel's "Summarize value field by" list (and its Custom subtotal functions), in Excel's order —
// the values match the backend's DimensionAggregationType (2026-10-15: Product … Varp added).
export const SUMMARY_FUNCTIONS = ["Sum", "Count", "Average", "Max", "Min", "Product", "CountNumbers", "StdDev", "StdDevp", "Var", "Varp"];

// The word Excel shows for a function ("Count Numbers"), mirroring DimensionAggregationNames.Display.
export function aggregationLabel(fn) {
  return fn === "CountNumbers" ? "Count Numbers" : fn;
}
