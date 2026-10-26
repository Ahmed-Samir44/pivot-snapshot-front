import { describe, expect, it } from "vitest";
import { fieldCaption, fieldCaptions, formatDimensionName } from "./cubeMeta";
import { MEASURES_PLACEHOLDER } from "./measuresPlacement";

// Excel labels an OLAP field by its own caption (2026-10-26).
describe("fieldCaption / fieldCaptions", () => {
  it("gives a field's own caption from a raw id or a Table.Field name", () => {
    expect(fieldCaption("[Dim Business Unit].[BU Short Name].[BU Short Name]")).toBe("BU Short Name");
    expect(fieldCaption("[Dim Date - Service Date].[FormattedDateHierarchy].[Month]")).toBe("Month");
    expect(fieldCaption("Business Unit.Country Name")).toBe("Country Name");
    expect(fieldCaption(MEASURES_PLACEHOLDER)).toBe("Σ Values");
  });

  it("keeps the full name only for fields whose captions would collide in the same PivotTable", () => {
    const serviceDate = "[Dim Date - Service Date].[Date].[Date]";
    const dischargeDate = "[Dim Date - Discharge Date].[Date].[Date]";
    const bu = "[Dim Business Unit].[BU Short Name].[BU Short Name]";
    const captionOf = fieldCaptions([serviceDate, dischargeDate, bu]);
    expect(captionOf(bu)).toBe("BU Short Name");
    expect(captionOf(serviceDate)).toBe(formatDimensionName(serviceDate));
    expect(captionOf(dischargeDate)).toBe(formatDimensionName(dischargeDate));
    // The same field listed twice (Rows and a filter) is not a collision.
    expect(fieldCaptions([bu, bu])(bu)).toBe("BU Short Name");
  });
});
