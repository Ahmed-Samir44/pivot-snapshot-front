// Shared between PivotChart.jsx (the live SVG chart) and PivotBuilder.jsx (the "why no chart"
// message) — kept in its own file rather than exported from PivotChart.jsx directly, since mixing
// component and non-component exports in one file breaks React Fast Refresh.
import { measuresIndexOf } from "./measuresPlacement";

// Mirrors the backend's HtmlSnapshotRenderer/PivotChartRenderer exactly: every detail column is a
// series, so Σ Values can sit anywhere in Columns (2026-10-15 — it used to have to be innermost).
// With Σ Values in ROWS each row is a measure, not a category, so there's no chart.
export function measuresPlacementSupportsChart(rows) {
  return measuresIndexOf(rows ?? []) < 0;
}

// Only covers the one condition that's simple, common, and worth a specific message; rarer edge
// cases (all-zero/negative values, no measures at all) fall through to no message, same as
// before — those are self-evident from the table itself.
export function chartUnavailableReason(result, lastQuery) {
  if (!result || !lastQuery?.values?.length || result.columnHeaders.length === 0) {
    return null;
  }

  if (!measuresPlacementSupportsChart(lastQuery.rows)) {
    return "No chart is drawn while Σ Values is in Rows.";
  }

  return null;
}
