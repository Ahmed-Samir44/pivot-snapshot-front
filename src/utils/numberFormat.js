// Mirrors the backend's NumberFormatter.cs exactly, so the live grid and the saved HTML snapshot
// always display a value field's numbers identically. Kept deliberately simple (three format
// types) rather than a full Excel custom-format mini-language — this is display-only and never
// touches the underlying number used for sorting/totals/etc.
const MIN_DECIMAL_PLACES = 0;
const MAX_DECIMAL_PLACES = 10;

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function withThousands(value, decimals) {
  return value.toLocaleString("en-US", { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

export function formatNumber(value, format) {
  if (value === null || value === undefined) {
    return "";
  }

  const type = format?.type ?? "General";
  const decimals = clamp(format?.decimalPlaces ?? 2, MIN_DECIMAL_PLACES, MAX_DECIMAL_PLACES);

  if (type === "Number") {
    return withThousands(value, decimals);
  }
  if (type === "Currency") {
    return `${format?.currencySymbol ?? "EGP"} ${withThousands(value, decimals)}`;
  }
  if (type === "Percentage") {
    // No *100 here — mirrors the backend's NumberFormatter exactly: ShowValuesAs percent variants
    // already produce percentage-scale numbers (38.43, not 0.3843).
    return `${withThousands(value, decimals)}%`;
  }

  // General: trims trailing zeros (up to 2 decimals) instead of always showing a fixed count.
  return value.toLocaleString("en-US", { maximumFractionDigits: 2 });
}

// Given a column index and the request's Values list, identifies which PivotValueField that
// column belongs to — the same "measures are the fastest-varying factor" invariant the backend's
// GrandTotalsProcessor/ShowValuesAsProcessor/NumberFormatter all rely on.
export function formatForColumn(columnIndex, values) {
  if (!values || values.length === 0) {
    return null;
  }
  return values[columnIndex % values.length]?.format ?? null;
}
