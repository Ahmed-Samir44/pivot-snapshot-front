import { describe, expect, it } from "vitest";
import { buildSlicerProbeRequest, itemsWithData } from "./slicerProbe";

// Excel greys out slicer items with no data under the other filters (2026-10-26).
describe("slicer probe", () => {
  const request = {
    rows: ["[BU]"],
    columns: ["[Country]", "[Measures]"],
    values: [{ field: "Revenue", showValuesAs: "Normal", format: { type: "Number" } }],
    filters: [
      { field: "[Country]", mode: "Members", includedMembers: ["Egypt"] },
      { field: "[Month]", mode: "Members", includedMembers: ["October 2026"] },
    ],
    calculatedFields: [],
    allowMultipleFiltersPerField: false,
  };

  it("asks for the slicer's field under every OTHER filter, with one plain value and no totals", () => {
    const probe = buildSlicerProbeRequest(request, "[Country]");
    expect(probe.rows).toEqual(["[Country]"]);
    expect(probe.values).toEqual([{ field: "Revenue", showValuesAs: "Normal" }]);
    expect(probe.filters.map((f) => f.field)).toEqual(["[Month]"]);
    expect(probe.grandTotals).toBe("Off");
    expect(probe.subtotals).toBe("Off");
  });

  it("has nothing to ask without a value field", () => {
    expect(buildSlicerProbeRequest({ ...request, values: [] }, "[Country]")).toBeNull();
  });

  it("counts an item as having data only when it came back with a number", () => {
    const result = {
      rowHeaders: [{ labels: ["Egypt"] }, { labels: ["KSA"] }, { labels: ["Total"], isTotal: true }],
      cells: [[5], [null], [5]],
    };
    expect([...itemsWithData(result)]).toEqual(["Egypt"]);
  });
});
