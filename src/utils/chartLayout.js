// Shared between PivotChart.jsx (the live SVG chart) and PivotBuilder.jsx (the "why no chart"
// message) — kept in its own file rather than exported from PivotChart.jsx directly, since mixing
// component and non-component exports in one file breaks React Fast Refresh.
import { measuresIndexOf } from "./measuresPlacement";

// Mirrors the backend's HtmlSnapshotRenderer/PivotChartRenderer exactly: a chart needs Σ Values to
// be in Columns AND innermost there (a fixed, contiguous stride — see PivotChartRenderer.Render's
// own WHY comment for why that's narrower than the general positional model everything else now
// supports).
export function measuresPlacementSupportsChart(rows, columns) {
  const rowsMeasuresIndex = measuresIndexOf(rows ?? []);
  const columnsMeasuresIndex = measuresIndexOf(columns ?? []);
  return rowsMeasuresIndex < 0 && columnsMeasuresIndex === (columns?.length ?? 0) - 1;
}

// Only covers the one condition that's simple, common, and worth a specific message; rarer edge
// cases (all-zero/negative values, no measures at all) fall through to no message, same as
// before — those are self-evident from the table itself.
export function chartUnavailableReason(result, lastQuery) {
  if (!result || !lastQuery?.values?.length || result.columnHeaders.length === 0) {
    return null;
  }

  if (!measuresPlacementSupportsChart(lastQuery.rows, lastQuery.columns)) {
    return "No chart is drawn unless Σ Values is the innermost field in Columns.";
  }

  return null;
}
