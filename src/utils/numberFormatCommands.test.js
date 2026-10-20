import { describe, expect, it } from "vitest";
import { accountingStyle, changeDecimals, commaStyle, percentStyle, sharedFormatType, withFormatType } from "./numberFormatCommands";
import { formatNumber } from "./numberFormat";

describe("Home › Number commands", () => {
  it("Increase Decimal from General switches to Number with one more decimal than shown", () => {
    const next = changeDecimals(undefined, +1, 2);
    expect(next).toMatchObject({ type: "Number", decimalPlaces: 3 });
    expect(formatNumber(4765.41, next)).toBe("4,765.410");
  });

  it("Decrease Decimal goes down to whole numbers and stops at 0", () => {
    let format = { type: "Number", decimalPlaces: 1 };
    format = changeDecimals(format, -1);
    expect(formatNumber(4765.41, format)).toBe("4,765");
    expect(changeDecimals(format, -1).decimalPlaces).toBe(0);
  });

  it("leaves a format without decimals alone", () => {
    expect(changeDecimals({ type: "Date" }, +1)).toMatchObject({ type: "Date" });
  });

  it("$, % and , set Excel's styles", () => {
    expect(formatNumber(1234.5, accountingStyle(undefined, "EGP"))).toBe("EGP 1,234.50");
    expect(formatNumber(0.384, percentStyle(undefined))).toBe("38%");
    expect(formatNumber(-1234.5, commaStyle({ type: "Currency", currencySymbol: "$" }))).toBe("(1,234.50)");
  });

  it("picking a category uses Excel's 2 decimals", () => {
    expect(withFormatType({ type: "Number", decimalPlaces: 0 }, "Currency")).toMatchObject({ type: "Currency", decimalPlaces: 2 });
  });

  it("the dropdown shows a type only when every selected field shares it", () => {
    expect(sharedFormatType([{ type: "Number" }, { type: "Number" }])).toBe("Number");
    expect(sharedFormatType([{ type: "Number" }, undefined])).toBeNull();
    expect(sharedFormatType([undefined])).toBe("General");
  });
});
