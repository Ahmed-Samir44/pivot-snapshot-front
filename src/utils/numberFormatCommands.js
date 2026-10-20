// Excel's Home › Number group (2026-10-20, "a way to select columns or some data and change the
// decimal point"), as plain functions from a value field's current format to its next one. A format
// belongs to a whole value field (as in an Excel PivotTable), so the ribbon applies these to every
// value field the selection covers.

const BASE = { type: "General", decimalPlaces: 2, currencySymbol: "" };
const MAX_DECIMALS = 10;
const clampDecimals = (n) => Math.min(MAX_DECIMALS, Math.max(0, n));

// The Number Format dropdown's choices, in Excel's order.
export const NUMBER_FORMAT_CHOICES = [
  { type: "General", label: "General" },
  { type: "Number", label: "Number" },
  { type: "Currency", label: "Currency" },
  { type: "Accounting", label: "Accounting" },
  { type: "Date", label: "Short Date" },
  { type: "Time", label: "Time" },
  { type: "Percentage", label: "Percentage" },
  { type: "Fraction", label: "Fraction" },
  { type: "Scientific", label: "Scientific" },
  { type: "Text", label: "Text" },
];

// Formats with a decimal count Increase/Decrease Decimal can change.
const HAS_DECIMALS = new Set(["Number", "Currency", "Accounting", "Percentage", "Scientific"]);

// Picking a category from the dropdown: Excel's defaults for it (2 decimals), keeping the symbol.
export function withFormatType(format, type) {
  const current = { ...BASE, ...format };
  return { ...current, type, decimalPlaces: HAS_DECIMALS.has(type) ? 2 : current.decimalPlaces };
}

// The $ button: Accounting, with the chosen currency symbol (or the field's own, or none).
export function accountingStyle(format, symbol) {
  const current = { ...BASE, ...format };
  return { ...current, type: "Accounting", currencySymbol: symbol ?? current.currencySymbol ?? "", decimalPlaces: current.type === "Accounting" ? current.decimalPlaces : 2 };
}

// The % button: Percent Style, no decimals.
export function percentStyle(format) {
  return { ...BASE, ...format, type: "Percentage", decimalPlaces: 0 };
}

// The , button: Comma Style — Accounting without a currency symbol, 2 decimals.
export function commaStyle(format) {
  return { ...BASE, ...format, type: "Accounting", currencySymbol: "", decimalPlaces: 2 };
}

// Increase (+1) / Decrease (-1) Decimal. From General it switches to Number, starting from the
// decimals the numbers show now (shownDecimals), like Excel; a format without decimals is unchanged.
export function changeDecimals(format, delta, shownDecimals = 2) {
  const current = { ...BASE, ...format };
  if (current.type === "General") return { ...current, type: "Number", decimalPlaces: clampDecimals(shownDecimals + delta) };
  if (!HAS_DECIMALS.has(current.type)) return current;
  return { ...current, decimalPlaces: clampDecimals((current.decimalPlaces ?? 2) + delta) };
}

// The one format type every selected field shares (the dropdown shows it), else null.
export function sharedFormatType(formats) {
  const types = [...new Set(formats.map((f) => f?.type ?? "General"))];
  return types.length === 1 ? types[0] : null;
}
