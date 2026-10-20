// The Show Values As modes whose result is a percentage.
export const PERCENT_SHOW_VALUES_AS = new Set([
  "PercentOfGrandTotal",
  "PercentOfRowTotal",
  "PercentOfColumnTotal",
  "PercentOfParentRow",
  "PercentOfParentColumn",
  "PercentDifferenceFrom",
  "PercentRunningTotal",
  "PercentOf",
  "PercentOfParentTotal",
]);

// The value field's format when its Show Values As changes: a % mode turns the untouched General
// format into Percentage, and leaving a % mode for a non-% one turns it back (2026-10-20, found by
// the feature test: No Calculation after % of Grand Total showed 24,947,704,322.00%). A format the
// user picked themselves is left alone. Returns null when nothing changes.
export function formatForShowValuesAs(format, previousMode, nextMode) {
  const current = format ?? { type: "General", decimalPlaces: 2, currencySymbol: "" };
  if (PERCENT_SHOW_VALUES_AS.has(nextMode) && current.type === "General") return { ...current, type: "Percentage" };
  if (!PERCENT_SHOW_VALUES_AS.has(nextMode) && PERCENT_SHOW_VALUES_AS.has(previousMode ?? "Normal") && current.type === "Percentage") return { ...current, type: "General" };
  return null;
}
