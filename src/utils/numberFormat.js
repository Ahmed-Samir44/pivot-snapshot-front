// Mirrors the backend's NumberFormatter.cs exactly, so the live grid and the saved HTML snapshot
// always display a value field's numbers identically. Covers all 11 of Excel's own Format Cells
// categories, including Special and Custom (see formatAsSpecial/formatAsCustom below, which mirror
// NumberFormatter.cs's own FormatAsSpecial/FormatAsCustom scope and limitations exactly). Display-
// only, never touches the underlying number used for sorting/totals.
const MIN_DECIMAL_PLACES = 0;
const MAX_DECIMAL_PLACES = 10;

// Excel's own date/time serial epoch — see the backend's identical constant for why 1899-12-30,
// not 1899-12-31 (Excel's historical 1900-leap-year bug).
const EXCEL_EPOCH_UTC_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 24 * 60 * 60 * 1000;
const MAX_FRACTION_DENOMINATOR = 99;

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function withThousands(value, decimals, useGrouping = true) {
  return value.toLocaleString("en-US", { minimumFractionDigits: decimals, maximumFractionDigits: decimals, useGrouping });
}

// Excel's Number/Currency "Negative numbers" list (2026-10-21): -1,234.10 (Minus), red 1,234.10
// (Red), (1,234.10) (Parentheses), red (1,234.10) (RedParentheses). Red drops the minus sign, as in
// Excel; the colour itself is the grid's job (isRedNegative).
export const NEGATIVE_NUMBER_STYLES = ["Minus", "Red", "Parentheses", "RedParentheses"];
const supportsNegativeStyle = (type) => type === "Number" || type === "Currency";

export function isRedNegative(value, format) {
  return typeof value === "number" && value < 0 && supportsNegativeStyle(format?.type) && (format?.negativeNumbers ?? "Minus").startsWith("Red");
}

function withNegativeStyle(value, format, render) {
  const style = format?.negativeNumbers ?? "Minus";
  if (value >= 0 || style === "Minus") return render(value);
  if (style === "Red") return render(Math.abs(value));
  return `(${render(Math.abs(value))})`;
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function formatAsExcelSerialDate(value) {
  const date = new Date(EXCEL_EPOCH_UTC_MS + value * MS_PER_DAY);
  return `${date.getUTCFullYear()}-${pad2(date.getUTCMonth() + 1)}-${pad2(date.getUTCDate())}`;
}

function formatAsExcelSerialTime(value) {
  // Only the FRACTIONAL part is time-of-day (0.5 = 12:00:00) — the whole-number (date) part is
  // ignored, same as Excel's own "Time" category applied to a value that also has a date part.
  const fraction = value - Math.floor(value);
  const totalSeconds = Math.round(fraction * 86400);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  return `${pad2(hours)}:${pad2(minutes)}:${pad2(seconds)}`;
}

// Simplest continued-fraction reduction with a capped denominator — matches the backend's own
// FormatAsFraction exactly (see its WHY comment for why this, not every one of Excel's half-dozen
// fraction sub-styles).
function formatAsFraction(value) {
  const isNegative = value < 0;
  const whole = Math.floor(Math.abs(value));
  const fractional = Math.abs(value) - whole;

  if (fractional < 1 / (MAX_FRACTION_DENOMINATOR * 2)) {
    return isNegative && whole > 0 ? `-${whole}` : String(whole);
  }

  let bestNumerator = 1;
  let bestDenominator = 1;
  let bestError = Number.MAX_VALUE;
  for (let denominator = 1; denominator <= MAX_FRACTION_DENOMINATOR; denominator++) {
    const numerator = Math.round(fractional * denominator);
    if (numerator === 0) continue;
    const error = Math.abs(fractional - numerator / denominator);
    if (error < bestError) {
      bestError = error;
      bestNumerator = numerator;
      bestDenominator = denominator;
    }
  }

  const sign = isNegative ? "-" : "";
  return whole > 0 ? `${sign}${whole} ${bestNumerator}/${bestDenominator}` : `${sign}${bestNumerator}/${bestDenominator}`;
}

// Mirrors NumberFormatter.cs's FormatAsSpecial exactly: a fixed digit-mask, value treated as a
// plain digit sequence (left-padded with zeros, truncated to the rightmost digits if too long),
// not an amount.
const SPECIAL_MASKS = {
  ZipCode: "00000",
  ZipCodePlus4: "00000-0000",
  PhoneNumber: "(000) 000-0000",
  SocialSecurityNumber: "000-00-0000",
};

function formatAsSpecial(value, specialType) {
  const mask = SPECIAL_MASKS[specialType] ?? SPECIAL_MASKS.ZipCode;
  const digitCount = mask.split("").filter((c) => c === "0").length;
  let digits = String(Math.abs(Math.round(value))).padStart(digitCount, "0");
  if (digits.length > digitCount) {
    digits = digits.slice(digits.length - digitCount);
  }
  let digitIndex = 0;
  return mask
    .split("")
    .map((c) => (c === "0" ? digits[digitIndex++] : c))
    .join("");
}

// Strips quoted-text delimiters (keeping their contents), resolves backslash escapes, and drops
// [bracketed] color/condition tags entirely — same scope as the backend's StripCustomFormatLiterals.
function stripCustomFormatLiterals(s) {
  let result = "";
  for (let i = 0; i < s.length; i++) {
    if (s[i] === '"') continue;
    if (s[i] === "\\" && i + 1 < s.length) {
      result += s[i + 1];
      i++;
      continue;
    }
    if (s[i] === "[") {
      const close = s.indexOf("]", i);
      if (close > i) {
        i = close;
        continue;
      }
    }
    result += s[i];
  }
  return result;
}

function matchToken(section, index, token) {
  return section.slice(index, index + token.length).toLowerCase() === token;
}

function renderCustomDateTime(section, value) {
  const date = new Date(EXCEL_EPOCH_UTC_MS + value * MS_PER_DAY);
  const lower = section.toLowerCase();
  const minutesAreTime = lower.includes("hh") || lower.includes("ss");
  let result = "";
  for (let i = 0; i < section.length; ) {
    if (section[i] === '"') {
      let close = section.indexOf('"', i + 1);
      if (close < 0) close = section.length - 1;
      result += section.slice(i + 1, close);
      i = close + 1;
      continue;
    }
    if (section[i] === "[") {
      const close = section.indexOf("]", i);
      i = close < 0 ? section.length : close + 1;
      continue;
    }
    if (matchToken(section, i, "yyyy")) {
      result += date.getUTCFullYear();
      i += 4;
      continue;
    }
    if (matchToken(section, i, "yy")) {
      result += pad2(date.getUTCFullYear() % 100);
      i += 2;
      continue;
    }
    if (matchToken(section, i, "mm")) {
      result += minutesAreTime ? pad2(date.getUTCMinutes()) : pad2(date.getUTCMonth() + 1);
      i += 2;
      continue;
    }
    if (matchToken(section, i, "m")) {
      result += minutesAreTime ? date.getUTCMinutes() : date.getUTCMonth() + 1;
      i += 1;
      continue;
    }
    if (matchToken(section, i, "dd")) {
      result += pad2(date.getUTCDate());
      i += 2;
      continue;
    }
    if (matchToken(section, i, "d")) {
      result += date.getUTCDate();
      i += 1;
      continue;
    }
    if (matchToken(section, i, "hh")) {
      result += pad2(date.getUTCHours());
      i += 2;
      continue;
    }
    if (matchToken(section, i, "h")) {
      result += date.getUTCHours();
      i += 1;
      continue;
    }
    if (matchToken(section, i, "ss")) {
      result += pad2(date.getUTCSeconds());
      i += 2;
      continue;
    }
    if (matchToken(section, i, "s")) {
      result += date.getUTCSeconds();
      i += 1;
      continue;
    }
    result += section[i];
    i += 1;
  }
  return result;
}

// Mirrors NumberFormatter.cs's FormatAsCustom exactly — same best-effort subset (up to 2 sections,
// digit placeholders with grouping/decimals, literal '%' scaling by 100, quoted/escaped/bracketed
// literals, and the same simplified date/time token detection) and the same deliberate omissions
// (no AM/PM, no elapsed time, no fractions, no locale-specific symbols beyond what's typed).
function formatAsCustom(value, code) {
  if (!code || !code.trim()) {
    return value.toLocaleString("en-US", { maximumFractionDigits: 2 });
  }

  const sections = code.split(";");
  const isNegative = value < 0;
  let section;
  let magnitude;
  if (value === 0 && sections.length >= 3) {
    section = sections[2];
    magnitude = 0;
  } else if (isNegative && sections.length >= 2) {
    section = sections[1];
    magnitude = Math.abs(value);
  } else {
    section = sections[0];
    magnitude = Math.abs(value);
  }
  const forceNegativeSign = isNegative && sections.length < 2;

  const lower = section.toLowerCase();
  const looksLikeDateTime = ["yyyy", "yy", "dd", "hh", "ss"].some((t) => lower.includes(t));
  if (looksLikeDateTime) {
    return renderCustomDateTime(section, magnitude);
  }

  const hasPercent = section.includes("%");
  if (hasPercent) {
    magnitude *= 100;
  }

  let skeletonStart = -1;
  let skeletonEnd = -1;
  for (let i = 0; i < section.length; i++) {
    if (["0", "#", ",", "."].includes(section[i])) {
      if (skeletonStart === -1) skeletonStart = i;
      skeletonEnd = i;
    }
  }

  if (skeletonStart === -1) {
    return stripCustomFormatLiterals(section);
  }

  const skeleton = section.slice(skeletonStart, skeletonEnd + 1);
  const dotIndex = skeleton.indexOf(".");
  const decimalPlaces = dotIndex >= 0 ? skeleton.length - dotIndex - 1 : 0;
  const integerPart = dotIndex >= 0 ? skeleton.slice(0, dotIndex) : skeleton;
  const minIntegerDigits = Math.max(1, integerPart.split("").filter((c) => c === "0").length);
  const useGrouping = integerPart.includes(",");

  const formattedNumber = magnitude.toLocaleString("en-US", {
    useGrouping,
    minimumIntegerDigits: minIntegerDigits,
    minimumFractionDigits: decimalPlaces,
    maximumFractionDigits: decimalPlaces,
  });

  const prefix = stripCustomFormatLiterals(section.slice(0, skeletonStart));
  const suffix = stripCustomFormatLiterals(section.slice(skeletonEnd + 1));
  const sign = forceNegativeSign ? "-" : "";
  return `${sign}${prefix}${formattedNumber}${suffix}`;
}

export function formatNumber(value, format) {
  // Never the literal text "NaN"/"∞". A value cell's error (NaN, since 2026-10-18) is shown by
  // pivotOptions.js's cellText before this runs; anything else non-finite stays blank.
  if (value === null || value === undefined || typeof value === "number" && !Number.isFinite(value)) {
    return "";
  }

  const type = format?.type ?? "General";
  const decimals = clamp(format?.decimalPlaces ?? 2, MIN_DECIMAL_PLACES, MAX_DECIMAL_PLACES);
  // No currency is assumed by default (mirrors the backend's NumberFormat.CurrencySymbol, which
  // defaults to "" — a specific currency is a per-report choice only the user makes, not something
  // this app assumes; an earlier "EGP" default was hardcoded here without ever having been asked
  // for, flagged live 2026-10-06). Empty stays empty — no stray leading space before the number.
  const symbol = format?.currencySymbol ?? "";
  const currencyPrefix = symbol ? `${symbol} ` : "";

  if (type === "Number") {
    // "Use 1000 Separator (,)" — on unless turned off (formats saved before it existed had it on).
    const grouping = format?.useThousandsSeparator !== false;
    return withNegativeStyle(value, format, (v) => withThousands(v, decimals, grouping));
  }
  if (type === "Currency") {
    return withNegativeStyle(value, format, (v) => `${currencyPrefix}${withThousands(v, decimals)}`);
  }
  if (type === "Percentage") {
    // *100 here — mirrors the backend's NumberFormatter exactly (fixed 2026-10-05, see its own WHY
    // comment): ShowValuesAs percent variants now produce a plain ratio (0.3843, not 38.43), same
    // as real Excel's own Percentage format and this app's Custom format's literal "%" handling.
    return `${withThousands(value * 100, decimals)}%`;
  }
  if (type === "Accounting") {
    // Negatives in parentheses instead of a leading minus sign; positives exactly like Currency;
    // zero shown as a plain dash, matching real Excel's own built-in Accounting format (its own
    // zero-specific format section never falls through to the positive one's "0.00" — mirrors the
    // same fix in the backend's NumberFormatter.cs, added 2026-10-06, full codebase review).
    if (value === 0) return `${currencyPrefix}-`;
    return value < 0 ? `(${currencyPrefix}${withThousands(Math.abs(value), decimals)})` : `${currencyPrefix}${withThousands(value, decimals)}`;
  }
  if (type === "Date") {
    return formatAsExcelSerialDate(value);
  }
  if (type === "Time") {
    return formatAsExcelSerialTime(value);
  }
  if (type === "Fraction") {
    return formatAsFraction(value);
  }
  if (type === "Scientific") {
    return value.toExponential(decimals).replace(/e([+-])(\d)$/, "E$10$2");
  }
  if (type === "Text") {
    // The raw number as plain digits, no thousands grouping — Excel's own "Text" category means
    // "don't apply any numeric formatting at all".
    return String(value);
  }
  if (type === "Special") {
    return formatAsSpecial(value, format?.specialType ?? "ZipCode");
  }
  if (type === "Custom") {
    return formatAsCustom(value, format?.customFormatCode ?? "");
  }

  // General: trims trailing zeros (up to 2 decimals) instead of always showing a fixed count.
  return value.toLocaleString("en-US", { maximumFractionDigits: 2 });
}

// measureIndex is the ALREADY-RESOLVED measure index (see utils/measuresPlacement.js's
// measureIndexOf) — this used to take a raw column/row index and derive the measure via `%
// values.length`, which only worked when measures were always the fastest-varying factor (fixed
// stride). Callers now resolve the real measure identity themselves from the header's own caption
// first, mirroring the backend's identical NumberFormatter.FormatFor change.
export function formatForColumn(measureIndex, values) {
  if (!values || measureIndex === null || measureIndex < 0 || measureIndex >= values.length) {
    return null;
  }
  return values[measureIndex]?.format ?? null;
}
