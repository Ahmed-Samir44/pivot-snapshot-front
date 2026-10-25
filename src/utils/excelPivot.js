import JSZip from "jszip";
import { MEASURES_PLACEHOLDER, measuresIndexOf, buildCaptionToIndex, measureIndexOf } from "./measuresPlacement";
import { filterValueSummary } from "./pivotOptions";
import { isErrorValue } from "./errorCells";

// A REAL Excel PivotTable in the exported file (2026-10-20, "why not like excel?"): Field List, the
// Row Labels dropdown, the PivotTable Analyze/Design tabs. Excel needs source rows to pivot, and the
// cube's result already holds them — every DETAIL cell (no subtotal/grand-total row or column) is one
// combination of the Rows/Columns items with its measure values. Those rows go on a "Data" sheet and
// the PivotTable sums them, which reproduces the app's numbers exactly for additive measures. Every
// total in the result is the cube's own aggregate, so the grand total is used to PROVE additivity:
// a measure whose detail cells don't add up to it (a distinct count, an average, a ratio) can't be
// re-summed by Excel, and the export falls back to plain values (as do Show Values As and
// drilldowns, which Excel can't rebuild from these rows).

const isGrandTotal = (h) => h.isTotal && (h.captionLevel === -1 || (h.captionLevel == null && h.labels.every((l) => l === "Total" || l === "Grand Total")));

// Labels per non-Σ field, by its position in the axis field list.
function axisFieldLabels(fields, labels) {
  const realLabels = labels.length === fields.length ? labels.filter((_, i) => fields[i] !== MEASURES_PLACEHOLDER) : labels.filter((l) => l !== "Σ Values");
  return fields.map((f, i) => ({ f, i })).filter(({ f }) => f !== MEASURES_PLACEHOLDER).map(({ f, i }, k) => ({ field: f, level: i, label: realLabels[k] ?? f }));
}

function uniqueNames(names) {
  const seen = new Map();
  return names.map((raw) => {
    const base = String(raw ?? "").trim() || "Field";
    const n = (seen.get(base.toLowerCase()) ?? 0) + 1;
    seen.set(base.toLowerCase(), n);
    return n === 1 ? base : `${base} (${n})`;
  });
}

// What the PivotTable is built from, or { ok: false, reason } when Excel couldn't reproduce the app's
// numbers from detail rows. Pure (no workbook) so it's unit-tested on its own.
export function buildPivotSource({ result, rows = [], columns = [], rowFieldLabels = [], columnFieldLabels = [], valueFields = [], filters = [], filterFieldLabels = {}, drilldown = false, drillSlot = -1, columnsDrillSlot = -1 }) {
  const no = (reason) => ({ ok: false, reason });
  if (!result || !valueFields.length) return no("there are no Values");
  if (drilldown || drillSlot >= 0 || columnsDrillSlot >= 0) return no("a drilldown hierarchy can't be rebuilt by Excel");
  if (valueFields.some((v) => v.showValuesAs && v.showValuesAs !== "Normal")) return no("Show Values As can't be rebuilt from detail rows");

  const { rowHeaders, columnHeaders, cells } = result;
  const rowsMi = measuresIndexOf(rows);
  const colsMi = measuresIndexOf(columns);
  const measuresInRows = rowsMi >= 0;
  const captionToIndex = buildCaptionToIndex(measuresInRows ? rowHeaders : columnHeaders, measuresInRows ? rowsMi : colsMi, valueFields.length);
  const measureOf = (rowHeader, colHeader) =>
    measuresInRows ? measureIndexOf(rowHeader, rowsMi, captionToIndex) : measureIndexOf(colHeader, colsMi, captionToIndex);

  const rowFields = axisFieldLabels(rows, rowFieldLabels);
  const colFields = axisFieldLabels(columns, columnFieldLabels);
  // Report Filter fields that narrow the data become page fields holding their one caption (the
  // member, or "(Multiple Items)"); a field left at (All) filters nothing and has no value to hold.
  const pageFilters = filters.filter((f) => f.field && !rows.includes(f.field) && !columns.includes(f.field) && filterValueSummary(f) !== "(All)");

  // One record per Rows × Columns item combination, in the result's own order.
  const records = new Map();
  rowHeaders.forEach((rowHeader, r) => {
    if (rowHeader.isTotal) return;
    columnHeaders.forEach((colHeader, c) => {
      if (colHeader.isTotal) return;
      const m = measureOf(rowHeader, colHeader);
      if (m === null) return;
      const rowValues = rowFields.map(({ level }) => rowHeader.labels[level] ?? "");
      const colValues = colFields.map(({ level }) => colHeader.labels[level] ?? "");
      const key = JSON.stringify([rowValues, colValues]);
      if (!records.has(key)) records.set(key, { rowValues, colValues, measures: valueFields.map(() => null) });
      const value = cells[r]?.[c];
      if (typeof value === "number" && !isErrorValue(value) && Number.isFinite(value)) records.get(key).measures[m] = value;
    });
  });
  const kept = [...records.values()].filter((rec) => rec.measures.some((v) => v !== null));
  if (kept.length === 0) return no("there's no data");

  // The cube's grand total for each measure, when the result has one — the detail cells must add up
  // to it, or Excel's re-summed totals would disagree with the app.
  const grandRows = rowFields.length > 0 ? rowHeaders.filter(isGrandTotal) : rowHeaders.filter((h) => !h.isTotal);
  const grandCols = colFields.length > 0 ? columnHeaders.filter(isGrandTotal) : columnHeaders.filter((h) => !h.isTotal);
  const grandOf = (m) => {
    for (const rh of grandRows) {
      for (const ch of grandCols) {
        if (measureOf(rh, ch) !== m) continue;
        const v = cells[rowHeaders.indexOf(rh)]?.[columnHeaders.indexOf(ch)];
        return typeof v === "number" && Number.isFinite(v) ? v : null;
      }
    }
    return null;
  };
  const measureNames = valueFields.map((v, m) => {
    const header = (measuresInRows ? rowHeaders : columnHeaders).find((h) => captionToIndex.get(h.labels[measuresInRows ? rowsMi : colsMi]) === m);
    return v.customName || header?.labels[measuresInRows ? rowsMi : colsMi] || v.field;
  });
  let verified = true;
  for (let m = 0; m < valueFields.length; m++) {
    const grand = grandOf(m);
    if (grand === null) {
      verified = false;
      continue;
    }
    const sum = kept.reduce((acc, rec) => acc + (rec.measures[m] ?? 0), 0);
    if (Math.abs(sum - grand) > 1e-6 * Math.max(1, Math.abs(grand))) {
      return no(`"${measureNames[m]}" doesn't add up across rows (a distinct count, average or ratio), so Excel would total it differently`);
    }
  }

  const names = uniqueNames([
    ...pageFilters.map((f) => filterFieldLabels[f.field] ?? f.field),
    ...rowFields.map((f) => f.label),
    ...colFields.map((f) => f.label),
    ...measureNames,
  ]);
  let n = 0;
  // Each field's items in the app's own header order (its A to Z, Z to A, by-value or manual sort),
  // keeping only items that carry data. Taking them in order of first appearance across the
  // records instead put a column item first just because the FIRST ROW had it — "Kingdom of Saudi
  // Arabia" before "Egypt" when AKW (KSA only) was the first row (found in Excel, 2026-10-25).
  const headerOrder = (headers, level, present) => [...new Set(headers.filter((h) => !h.isTotal).map((h) => h.labels[level] ?? ""))].filter((v) => present.has(v));
  const itemsOf = (headers, axisFields, k, valuesOf) => headerOrder(headers, axisFields[k].level, new Set(kept.map((rec) => valuesOf(rec)[k])));
  const fields = [
    ...pageFilters.map((f) => ({ kind: "page", name: names[n++], items: [filterValueSummary(f)] })),
    ...rowFields.map((_, k) => ({ kind: "row", name: names[n++], items: itemsOf(rowHeaders, rowFields, k, (rec) => rec.rowValues) })),
    ...colFields.map((_, k) => ({ kind: "col", name: names[n++], items: itemsOf(columnHeaders, colFields, k, (rec) => rec.colValues) })),
    ...valueFields.map((v, m) => ({ kind: "data", name: names[n++], measureIndex: m, caption: measureNames[m] })),
  ];
  const data = kept.map((rec) => [...pageFilters.map((f) => filterValueSummary(f)), ...rec.rowValues, ...rec.colValues, ...rec.measures]);

  return {
    ok: true,
    verified,
    fields,
    data,
    measuresInRows,
    // Where Σ Values sits among the real fields of its axis (Excel's "Values" field, x = -2).
    measuresPosition: measuresInRows ? rows.slice(0, rowsMi).filter((f) => f !== MEASURES_PLACEHOLDER).length : columns.slice(0, Math.max(colsMi, 0)).filter((f) => f !== MEASURES_PLACEHOLDER).length,
    // Excel's rowGrandTotals is the Grand Total ROW (totals of the row items), colGrandTotals the
    // Grand Total COLUMN — checked in Excel (rowGrandTotals="0" removed the bottom row).
    rowGrandTotals: rowHeaders.some(isGrandTotal),
    colGrandTotals: columnHeaders.some(isGrandTotal),
    rowSubtotals: rowHeaders.some((h) => h.isTotal && !isGrandTotal(h)),
    colSubtotals: columnHeaders.some((h) => h.isTotal && !isGrandTotal(h)),
  };
}

const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_PIVOT_CACHE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
const REL_PIVOT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";

// Excel's built-in number formats a data field can point at without a styles.xml entry.
const BUILTIN_FORMATS = { General: 0, "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9, "0.00%": 10 };

function numFmtIdFor(code, stylesXml) {
  if (code in BUILTIN_FORMATS) return BUILTIN_FORMATS[code];
  for (const m of stylesXml.matchAll(/<numFmt\s+numFmtId="(\d+)"\s+formatCode="([^"]*)"/g)) {
    const decoded = m[2].replace(/&quot;/g, '"').replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
    if (decoded === code) return Number(m[1]);
  }
  return 0;
}

// The sheet's part path ("xl/worksheets/sheet1.xml") from its tab name.
function sheetPath(workbookXml, relsXml, name) {
  const sheet = [...workbookXml.matchAll(/<sheet\b[^>]*>/g)].map((m) => m[0]).find((tag) => tag.includes(`name="${esc(name)}"`));
  const rId = sheet?.match(/r:id="([^"]+)"/)?.[1];
  const target = [...relsXml.matchAll(/<Relationship\b[^>]*>/g)].map((m) => m[0]).find((tag) => tag.includes(`Id="${rId}"`))?.match(/Target="([^"]+)"/)?.[1];
  if (!target) throw new Error(`Sheet "${name}" not found in the workbook.`);
  return target.startsWith("/") ? target.slice(1) : `xl/${target}`;
}

function nextRelId(relsXml) {
  const ids = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
  return `rId${(ids.length ? Math.max(...ids) : 0) + 1}`;
}

// Adds the PivotTable to an ExcelJS-written .xlsx (ExcelJS can't write one): a pivot cache over the
// Data sheet's range and the table itself at `location` on the pivot sheet. refreshOnLoad makes
// Excel rebuild it from the Data sheet when the file opens, so no cache records are written; the
// pivot sheet carries the same values as plain cells meanwhile (what Protected View shows).
// spec: { source (buildPivotSource), name, pivotSheet, dataSheet, dataRef, location, layout,
// subtotalsAtBottom, numFmtCodes (per data field), pageOverThenDown, pageWrap }
export async function addPivotTableToXlsx(buffer, spec) {
  const zip = await JSZip.loadAsync(buffer);
  const read = (p) => zip.file(p)?.async("string");
  let workbookXml = await read("xl/workbook.xml");
  let workbookRels = await read("xl/_rels/workbook.xml.rels");
  let contentTypes = await read("[Content_Types].xml");
  const stylesXml = (await read("xl/styles.xml")) ?? "";
  const pivotSheetPath = sheetPath(workbookXml, workbookRels, spec.pivotSheet);

  const { source } = spec;
  const fieldIndex = (f) => source.fields.indexOf(f);
  const of = (kind) => source.fields.filter((f) => f.kind === kind);
  const dataFields = of("data");

  // --- pivot cache definition ---
  const cacheFields = source.fields.map((f, i) => {
    if (f.kind === "data") {
      const values = source.data.map((row) => row[i]).filter((v) => v !== null);
      const hasBlank = values.length < source.data.length;
      const ints = values.every((v) => Number.isInteger(v));
      const range = values.length ? ` minValue="${Math.min(...values)}" maxValue="${Math.max(...values)}"` : "";
      return `<cacheField name="${esc(f.name)}" numFmtId="0"><sharedItems${hasBlank ? "" : ' containsSemiMixedTypes="0"'} containsString="0" containsNumber="1"${ints ? ' containsInteger="1"' : ""}${hasBlank ? ' containsBlank="1"' : ""}${range}/></cacheField>`;
    }
    const hasBlank = f.items.includes("");
    const items = f.items.map((v) => (v === "" ? "<m/>" : `<s v="${esc(v)}"/>`)).join("");
    return `<cacheField name="${esc(f.name)}" numFmtId="0"><sharedItems${hasBlank ? ' containsBlank="1"' : ""} count="${f.items.length}">${items}</sharedItems></cacheField>`;
  });
  const cacheXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_REL}" refreshOnLoad="1" saveData="0" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="${source.data.length}">` +
    `<cacheSource type="worksheet"><worksheetSource ref="${spec.dataRef}" sheet="${esc(spec.dataSheet)}"/></cacheSource>` +
    `<cacheFields count="${cacheFields.length}">${cacheFields.join("")}</cacheFields></pivotCacheDefinition>`;

  // --- pivot table definition ---
  const tabular = spec.layout === "Tabular";
  const outline = spec.layout === "Outline";
  const fieldLayoutAttrs = tabular ? ' compact="0" outline="0"' : outline ? ' compact="0"' : "";
  const pivotFields = source.fields.map((f) => {
    if (f.kind === "data") return `<pivotField dataField="1"${fieldLayoutAttrs} showAll="0"/>`;
    const axis = { row: "axisRow", col: "axisCol", page: "axisPage" }[f.kind];
    const subtotals = f.kind === "row" ? source.rowSubtotals : f.kind === "col" ? source.colSubtotals : true;
    const items = f.items.map((_, x) => `<item x="${x}"/>`).join("") + (subtotals ? '<item t="default"/>' : "");
    const count = f.items.length + (subtotals ? 1 : 0);
    const subtotalAttrs = (subtotals ? "" : ' defaultSubtotal="0"') + (f.kind === "row" && spec.subtotalsAtBottom ? ' subtotalTop="0"' : "");
    return `<pivotField axis="${axis}"${fieldLayoutAttrs} showAll="0"${subtotalAttrs}><items count="${count}">${items}</items></pivotField>`;
  });
  // Σ Values ("Values", x = -2) sits among its axis's fields where it sat in the app — only when
  // there are 2+ data fields, like Excel.
  const axisFieldsXml = (kind, isMeasuresAxis) => {
    const xs = of(kind).map((f) => fieldIndex(f));
    if (isMeasuresAxis && dataFields.length > 1) xs.splice(Math.min(source.measuresPosition, xs.length), 0, -2);
    return xs.length ? `<${kind}Fields count="${xs.length}">${xs.map((x) => `<field x="${x}"/>`).join("")}</${kind}Fields>` : "";
  };
  const pageFields = of("page");
  const dataFieldsXml = dataFields
    .map((f, k) => {
      const shown = f.caption === f.name ? `${f.caption} ` : f.caption; // a data field may not reuse its source field's name
      return `<dataField name="${esc(shown)}" fld="${fieldIndex(f)}" baseField="0" baseItem="0" numFmtId="${numFmtIdFor(spec.numFmtCodes[k] ?? "General", stylesXml)}"/>`;
    })
    .join("");
  const tableXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<pivotTableDefinition xmlns="${NS_MAIN}" name="${esc(spec.name)}" cacheId="${spec.cacheId ?? 1}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0"` +
    (source.measuresInRows ? ' dataOnRows="1"' : "") +
    (tabular ? ' compact="0" compactData="0"' : ' outline="1" outlineData="1"') +
    (outline ? ' compact="0" compactData="0"' : "") +
    (source.rowGrandTotals ? "" : ' rowGrandTotals="0"') +
    (source.colGrandTotals ? "" : ' colGrandTotals="0"') +
    (spec.pageOverThenDown ? ' pageOverThenDown="1"' : "") +
    (spec.pageWrap ? ` pageWrap="${spec.pageWrap}"` : "") +
    ` multipleFieldFilters="0">` +
    `<location ref="${spec.location}" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"${pageFields.length ? ` rowPageCount="${pageFields.length}" colPageCount="1"` : ""}/>` +
    `<pivotFields count="${pivotFields.length}">${pivotFields.join("")}</pivotFields>` +
    axisFieldsXml("row", source.measuresInRows) +
    axisFieldsXml("col", !source.measuresInRows) +
    (pageFields.length ? `<pageFields count="${pageFields.length}">${pageFields.map((f) => `<pageField fld="${fieldIndex(f)}" item="0" hier="-1"/>`).join("")}</pageFields>` : "") +
    `<dataFields count="${dataFields.length}">${dataFieldsXml}</dataFields>` +
    `<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>` +
    `</pivotTableDefinition>`;

  // --- wire the parts together ---
  const cachePath = "xl/pivotCache/pivotCacheDefinition1.xml";
  const tablePath = "xl/pivotTables/pivotTable1.xml";
  zip.file(cachePath, cacheXml);
  zip.file(tablePath, tableXml);
  zip.file(
    "xl/pivotTables/_rels/pivotTable1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_PIVOT_CACHE}" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>`,
  );

  const cacheRelId = nextRelId(workbookRels);
  workbookRels = workbookRels.replace("</Relationships>", `<Relationship Id="${cacheRelId}" Type="${REL_PIVOT_CACHE}" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>`);
  // <pivotCaches> follows <calcPr> (or <definedNames>/<sheets>) in the workbook's element order.
  const pivotCaches = `<pivotCaches><pivotCache cacheId="${spec.cacheId ?? 1}" r:id="${cacheRelId}"/></pivotCaches>`;
  const anchor = workbookXml.match(/<calcPr\b[^>]*\/>|<\/calcPr>|<\/definedNames>|<\/sheets>/g)?.pop();
  workbookXml = anchor ? workbookXml.replace(anchor, `${anchor}${pivotCaches}`) : workbookXml.replace("</workbook>", `${pivotCaches}</workbook>`);
  if (!/xmlns:r=/.test(workbookXml.slice(0, workbookXml.indexOf(">", workbookXml.indexOf("<workbook"))))) {
    workbookXml = workbookXml.replace("<workbook ", `<workbook xmlns:r="${NS_REL}" `);
  }

  const sheetFile = pivotSheetPath.split("/").pop();
  const sheetRelsPath = pivotSheetPath.replace(sheetFile, `_rels/${sheetFile}.rels`);
  let sheetRels = (await read(sheetRelsPath)) ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  sheetRels = sheetRels.replace("</Relationships>", `<Relationship Id="${nextRelId(sheetRels)}" Type="${REL_PIVOT_TABLE}" Target="../pivotTables/pivotTable1.xml"/></Relationships>`);

  contentTypes = contentTypes.replace(
    "</Types>",
    `<Override PartName="/${cachePath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>` +
      `<Override PartName="/${tablePath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/></Types>`,
  );

  zip.file("xl/workbook.xml", workbookXml);
  zip.file("xl/_rels/workbook.xml.rels", workbookRels);
  zip.file(sheetRelsPath, sheetRels);
  zip.file("[Content_Types].xml", contentTypes);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
