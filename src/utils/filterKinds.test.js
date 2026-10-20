import { describe, expect, it } from "vitest";
import { filterKind, onePerField, setFieldFilter } from "./filterKinds";

const members = (field, ...m) => ({ field, mode: "Members", includedMembers: m });
const label = (field, text) => ({ field, mode: "LabelContains", includedMembers: [], labelText: text });
const top = (field) => ({ field, mode: "TopN", includedMembers: [], n: 5, byMeasureField: "Revenue" });

describe("filterKind", () => {
  it("puts manual, label/date and value filters in Excel's three kinds", () => {
    expect(filterKind(members("A"))).toBe(0);
    expect(filterKind({ field: "A", mode: "Exclude", includedMembers: [] })).toBe(0);
    expect(filterKind(label("A", "x"))).toBe(1);
    expect(filterKind({ field: "A", mode: "DatePeriod", labelText: "Today" })).toBe(1);
    expect(filterKind(top("A"))).toBe(2);
  });
});

describe("setFieldFilter", () => {
  it("replaces every filter on the field when only one is allowed, keeping its place", () => {
    const filters = [members("A", "x"), label("B", "y"), top("A")];
    const next = setFieldFilter(filters, label("A", "z"), false);
    expect(next).toEqual([label("A", "z"), label("B", "y")]);
  });

  it("replaces only the same kind when multiple filters per field are allowed", () => {
    const filters = [members("A", "x"), top("A")];
    const next = setFieldFilter(filters, label("A", "z"), true);
    expect(next).toEqual([members("A", "x"), top("A"), label("A", "z")]);
    expect(setFieldFilter(next, members("A", "q"), true)[0]).toEqual(members("A", "q"));
  });
});

describe("onePerField", () => {
  it("keeps each field's first filter in Excel's order (manual before label before value)", () => {
    const kept = onePerField([top("A"), label("A", "x"), members("A", "m"), top("B")]);
    expect(kept).toEqual([members("A", "m"), top("B")]);
  });
});
