import { describe, expect, it } from "vitest";
import { formatForShowValuesAs } from "./showValuesAsFormat";

describe("formatForShowValuesAs", () => {
  it("switches General to Percentage for a % mode, and back for No Calculation", () => {
    const pct = formatForShowValuesAs({ type: "General", decimalPlaces: 2 }, "Normal", "PercentOfGrandTotal");
    expect(pct.type).toBe("Percentage");
    expect(formatForShowValuesAs(pct, "PercentOfGrandTotal", "Normal").type).toBe("General");
  });

  it("leaves a format the user picked alone", () => {
    expect(formatForShowValuesAs({ type: "Number" }, "Normal", "PercentOfGrandTotal")).toBeNull();
    expect(formatForShowValuesAs({ type: "Percentage" }, "Normal", "Rank")).toBeNull();
  });
});
