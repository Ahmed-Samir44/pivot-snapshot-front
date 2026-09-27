// Shared between PivotChart.jsx (the live SVG chart) and PivotBuilder.jsx (the "why no chart"
// message) — kept in its own file rather than exported from PivotChart.jsx directly, since mixing
// component and non-component exports in one file breaks React Fast Refresh.

// Mirrors the backend's PivotChartRenderer.cs MaxBars exactly.
export const MAX_BARS = 50;

// Excel itself doesn't chart hundreds of categories either — this just makes the SAME cap the
// chart already enforces visible as an explicit reason instead of a silently empty space where a
// chart was expected. Only covers the two conditions that are simple, common, and worth a
// specific message; rarer edge cases (all-zero/negative values, no measures at all) fall through
// to no message, same as before — those are self-evident from the table itself.
export function chartUnavailableReason(result, lastQuery) {
  if (!result || !lastQuery?.values?.length || result.columnHeaders.length === 0) {
    return null;
  }

  if (lastQuery.valuesPlacement === "Rows") {
    return 'No chart is drawn while values are on Rows — see the "Values on" setting above.';
  }

  const categoryCount = result.rowHeaders.filter((r) => !r.isTotal).length;
  if (categoryCount > MAX_BARS) {
    return `No chart is drawn: this pivot has ${categoryCount} rows, more than the chart's limit of ${MAX_BARS}. Filter or group the Rows field(s) down to ${MAX_BARS} or fewer to see a chart.`;
  }

  return null;
}
