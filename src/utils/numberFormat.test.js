import { describe, expect, it } from "vitest";
import { formatNumber, isRedNegative } from "./numberFormat";

// Excel's Format Cells › Number: "Use 1000 Separator (,)" and the "Negative numbers" list (2026-10-21).
describe("formatNumber — Number/Currency options", () => {
  const number = (extra) => ({ type: "Number", decimalPlaces: 2, ...extra });

  it("keeps the separator on by default and drops it when unticked", () => {
    expect(formatNumber(1234.5, number())).toBe("1,234.50");
    expect(formatNumber(1234.5, number({ useThousandsSeparator: false }))).toBe("1234.50");
  });

  it("shows negatives the four ways Excel offers", () => {
    expect(formatNumber(-1234.1, number())).toBe("-1,234.10");
    expect(formatNumber(-1234.1, number({ negativeNumbers: "Red" }))).toBe("1,234.10");
    expect(formatNumber(-1234.1, number({ negativeNumbers: "Parentheses" }))).toBe("(1,234.10)");
    expect(formatNumber(-1234.1, number({ negativeNumbers: "RedParentheses" }))).toBe("(1,234.10)");
    // Positives are untouched by the negative style.
    expect(formatNumber(1234.1, number({ negativeNumbers: "Parentheses" }))).toBe("1,234.10");
  });

  it("applies the negative style to Currency too, with the symbol inside the parentheses", () => {
    expect(formatNumber(-5, { type: "Currency", decimalPlaces: 0, currencySymbol: "EGP", negativeNumbers: "Parentheses" })).toBe("(EGP 5)");
  });

  it("marks only negative values with a red style as red", () => {
    expect(isRedNegative(-1, number({ negativeNumbers: "Red" }))).toBe(true);
    expect(isRedNegative(-1, number({ negativeNumbers: "RedParentheses" }))).toBe(true);
    expect(isRedNegative(1, number({ negativeNumbers: "Red" }))).toBe(false);
    expect(isRedNegative(-1, number({ negativeNumbers: "Parentheses" }))).toBe(false);
    expect(isRedNegative(-1, { type: "Percentage", negativeNumbers: "Red" })).toBe(false);
  });
});
