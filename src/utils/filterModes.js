// Every rule-based filter mode, grouped the way Excel's Label Filters / Value Filters submenus list
// them (2026-10-20: the full set — it used to be Contains/Begins/Ends and Top N/>/</Between only).
// Mirrors the backend's FilterMode enum and FilterModes.IsLabel. "Top 10…" is ONE entry, like
// Excel's: its own dialog picks Top/Bottom and Items/Percent/Sum (TOP_10_MODES).

export const LABEL_FILTERS = [
  { value: "LabelEquals", label: "Equals…" },
  { value: "LabelDoesNotEqual", label: "Does Not Equal…" },
  { value: "LabelBeginsWith", label: "Begins With…" },
  { value: "LabelDoesNotBeginWith", label: "Does Not Begin With…" },
  { value: "LabelEndsWith", label: "Ends With…" },
  { value: "LabelDoesNotEndWith", label: "Does Not End With…" },
  { value: "LabelContains", label: "Contains…" },
  { value: "LabelDoesNotContain", label: "Does Not Contain…" },
  { value: "LabelGreaterThan", label: "Greater Than…" },
  { value: "LabelGreaterThanOrEqual", label: "Greater Than Or Equal To…" },
  { value: "LabelLessThan", label: "Less Than…" },
  { value: "LabelLessThanOrEqual", label: "Less Than Or Equal To…" },
  { value: "LabelBetween", label: "Between…" },
  { value: "LabelNotBetween", label: "Not Between…" },
];

export const VALUE_FILTERS = [
  { value: "ValueEquals", label: "Equals…" },
  { value: "ValueDoesNotEqual", label: "Does Not Equal…" },
  { value: "GreaterThan", label: "Greater Than…" },
  { value: "GreaterThanOrEqual", label: "Greater Than Or Equal To…" },
  { value: "LessThan", label: "Less Than…" },
  { value: "LessThanOrEqual", label: "Less Than Or Equal To…" },
  { value: "Between", label: "Between…" },
  { value: "NotBetween", label: "Not Between…" },
  { value: "TopN", label: "Top 10…" },
];

// Excel's Top 10 dialog: [Top|Bottom] [n] [Items|Percent|Sum].
export const TOP_10_MODES = {
  Top: { Items: "TopN", Percent: "TopPercent", Sum: "TopSum" },
  Bottom: { Items: "BottomN", Percent: "BottomPercent", Sum: "BottomSum" },
};
export const TOP_10_MODE_SET = new Set(Object.values(TOP_10_MODES).flatMap((m) => Object.values(m)));

const LABEL_MODE_SET = new Set(LABEL_FILTERS.map((f) => f.value));

export const isLabelMode = (mode) => LABEL_MODE_SET.has(mode);
export const isTop10Mode = (mode) => TOP_10_MODE_SET.has(mode);
export const isRangeMode = (mode) => mode === "Between" || mode === "NotBetween" || mode === "LabelBetween" || mode === "LabelNotBetween";
// A value filter that compares the measure with a number (not Top 10).
export const isValueCompareMode = (mode) => VALUE_FILTERS.some((f) => f.value === mode) && !isTop10Mode(mode);

// Where a Top 10 mode sits in the dialog: { direction: "Top"|"Bottom", unit: "Items"|"Percent"|"Sum" }.
export function top10Parts(mode) {
  for (const [direction, units] of Object.entries(TOP_10_MODES)) {
    for (const [unit, m] of Object.entries(units)) if (m === mode) return { direction, unit };
  }
  return { direction: "Top", unit: "Items" };
}

// Whether a rule-based filter has everything its mode needs (mirrors MdxPivotQueryBuilder.ValidateFilterShape).
export function isRuleFilterComplete(filter) {
  const mode = filter.mode;
  if (isLabelMode(mode)) return Boolean(filter.labelText) && (!isRangeMode(mode) || Boolean(filter.labelTextTo));
  if (mode === "TopN" || mode === "BottomN") return Boolean(filter.n > 0 && filter.byMeasureField);
  if (isTop10Mode(mode)) return filter.value != null && Boolean(filter.byMeasureField);
  if (isValueCompareMode(mode)) return filter.value != null && (!isRangeMode(mode) || filter.valueTo != null) && Boolean(filter.byMeasureField);
  return true;
}
