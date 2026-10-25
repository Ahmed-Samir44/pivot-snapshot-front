// The field's own name for a slicer/timeline title — "Country Name", not "Business Unit.Country
// Name" — the full name stays in the tooltip. Before the field list loads, displayName is the raw
// identifier ("[Dim X].[Country Name].[Country Name]"), so the brackets go too.
export function widgetTitle(displayName) {
  const name = String(displayName ?? "");
  const last = name.includes(".") ? name.slice(name.lastIndexOf(".") + 1) : name;
  return last.replace(/^\[|\]$/g, "");
}
