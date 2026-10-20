import { describe, expect, it } from "vitest";
import { DATE_PERIODS, isDateMode, isLabelMode, isRangeMode, isRuleFilterComplete, isTop10Mode, top10Parts, TOP_10_MODES } from "./filterModes";

describe("filter mode families", () => {
  it("tells label, date, Top 10 and range modes apart", () => {
    expect(isLabelMode("LabelContains")).toBe(true);
    expect(isLabelMode("GreaterThan")).toBe(false);
    expect(isDateMode("DatePeriod")).toBe(true);
    expect(isDateMode("LabelEquals")).toBe(false);
    expect(isTop10Mode("BottomPercent")).toBe(true);
    expect(isRangeMode("DateBetween")).toBe(true);
    expect(isRangeMode("LabelNotBetween")).toBe(true);
    expect(isRangeMode("GreaterThan")).toBe(false);
  });

  it("maps every Top 10 mode back to its direction and unit", () => {
    for (const [direction, units] of Object.entries(TOP_10_MODES)) {
      for (const [unit, mode] of Object.entries(units)) expect(top10Parts(mode)).toEqual({ direction, unit });
    }
  });

  it("offers every Excel date period, including each quarter and month", () => {
    const keys = DATE_PERIODS.map((p) => p.key);
    expect(keys).toEqual(expect.arrayContaining(["Today", "ThisMonth", "YearToDate", "Q1", "Q4", "M1", "M12"]));
    expect(new Set(keys).size).toBe(keys.length);
  });
});

describe("isRuleFilterComplete (mirrors the backend's ValidateFilterShape)", () => {
  it("needs text for a label filter, and both ends for a range", () => {
    expect(isRuleFilterComplete({ mode: "LabelContains", labelText: "" })).toBe(false);
    expect(isRuleFilterComplete({ mode: "LabelContains", labelText: "er" })).toBe(true);
    expect(isRuleFilterComplete({ mode: "LabelBetween", labelText: "a" })).toBe(false);
    expect(isRuleFilterComplete({ mode: "LabelBetween", labelText: "a", labelTextTo: "m" })).toBe(true);
  });

  it("needs a date for a date filter, and a period key for DatePeriod", () => {
    expect(isRuleFilterComplete({ mode: "DateBefore", labelText: "" })).toBe(false);
    expect(isRuleFilterComplete({ mode: "DateBetween", labelText: "2025-01-01", labelTextTo: "2025-01-31" })).toBe(true);
    expect(isRuleFilterComplete({ mode: "DatePeriod", labelText: "ThisMonth" })).toBe(true);
  });

  it("needs a number and a measure for value filters", () => {
    expect(isRuleFilterComplete({ mode: "TopN", n: 10, byMeasureField: null })).toBe(false);
    expect(isRuleFilterComplete({ mode: "TopN", n: 10, byMeasureField: "Revenue" })).toBe(true);
    expect(isRuleFilterComplete({ mode: "TopPercent", value: null, byMeasureField: "Revenue" })).toBe(false);
    expect(isRuleFilterComplete({ mode: "Between", value: 1, valueTo: null, byMeasureField: "Revenue" })).toBe(false);
    expect(isRuleFilterComplete({ mode: "GreaterThanOrEqual", value: 0, byMeasureField: "Revenue" })).toBe(true);
  });
});
