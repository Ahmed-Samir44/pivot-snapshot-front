// Real Excel's Design > Subtotals is not a separate report-wide flag sitting beside each field's own
// Field Settings > Subtotals — it's a bulk action that REWRITES every field's own setting:
// "Do Not Show Subtotals" sets every field to None, and "Show all Subtotals at Top/Bottom" sets every
// field to Automatic at that position. This app stores it as a report default (`subtotals`) plus
// per-field overrides (`fieldSubtotals`), so applying the bulk action means clearing the overrides it
// supersedes rather than leaving a stale override that contradicts what the user just picked.
// A Custom field keeps its functions and just moves to the new position (only None/Automatic are
// forced back to Automatic); "Do Not Show" drops every override, Custom included, since every field
// becomes None.
export function applyReportSubtotals(request, position) {
  if (position === "Off") {
    return { ...request, subtotals: "Off", fieldSubtotals: [] };
  }
  return {
    ...request,
    subtotals: position,
    // ?? []: a snapshot's request saved before fieldSubtotals existed (SnapshotHistory's PivotGrid).
    fieldSubtotals: (request.fieldSubtotals ?? [])
      .filter((f) => f.functions && f.functions.length > 0)
      .map((f) => ({ ...f, position: null })),
  };
}

// A field's EFFECTIVE subtotal position — its own override if it has one, else the report default.
// What Field Settings must show (Excel only ever shows a field's real state, never "inherit").
export function effectiveSubtotalPosition(override, reportDefault) {
  return override ?? reportDefault;
}
