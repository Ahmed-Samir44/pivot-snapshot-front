import { useEffect, useRef, useState } from "react";
import { FileSpreadsheet, Loader2, Plus } from "lucide-react";
import AppShell from "../components/layout/AppShell";
import FieldPicker from "../components/pivot/FieldPicker";
import PivotToolbar from "../components/pivot/PivotToolbar";
import PivotGrid from "../components/pivot/PivotGrid";
import PivotChart from "../components/pivot/PivotChart";
import { chartUnavailableReason } from "../utils/chartLayout";
import { queryPivot } from "../services/pivotApi";
import { getTables, getDimensionsForTable, getMeasures } from "../services/cubeMetaApi";
import { saveSnapshot, listDefinitions } from "../services/snapshotApi";
import { MEASURES_PLACEHOLDER, normalizeMeasuresPlacement, hideMeasuresLabelIfSingleValue } from "../utils/measuresPlacement";
import { readFieldDragPayload, FIELD_DRAG_MIME, removeFieldEverywhere } from "../utils/dragDrop";
import { isCompleteCalculatedField, isCompleteCalculatedItem } from "../utils/calculatedFields";

// This screen renders a live, interactive pivot (React state + re-fetch on every change).
// It is NOT the saved snapshot: "Save Snapshot" below freezes the CURRENT result (the one already
// on screen) into a static HTML snapshot — it never re-queries the cube, so what gets saved is
// guaranteed to match exactly what the user looked at before saving (see SnapshotStorageService).

// The expanded drilldown members that are actually VISIBLE — every ancestor of theirs is expanded
// too. A member with no recorded ancestry (state saved before 2026-10-14) counts as visible, which
// is exactly the old behavior.
function visibleDrilldownExpansions(drilldown) {
  const expanded = new Set(drilldown.expandedMemberUniqueNames);
  const ancestors = drilldown.expandedAncestors ?? {};
  return drilldown.expandedMemberUniqueNames.filter((u) => (ancestors[u] ?? []).every((a) => expanded.has(a)));
}

// Drilldown v3 (2026-10-14): real Excel expands a member PER OUTER ITEM, so with Rows fields before
// the drilldown slot, expansions live in `scopedExpansions` — { outerUniqueNames, memberUniqueName,
// ancestors } — instead of the context-free list. Same visibility rule as above, within one outer
// item: only expansions whose ancestors are all expanded under that same item are sent.
const outerKey = (outerUniqueNames) => outerUniqueNames.join("\u0001");
function visibleScopedExpansions(drilldown) {
  const scoped = drilldown.scopedExpansions ?? [];
  const expanded = new Set(scoped.map((e) => `${outerKey(e.outerUniqueNames)}\u0002${e.memberUniqueName}`));
  return scoped.filter((e) => (e.ancestors ?? []).every((a) => expanded.has(`${outerKey(e.outerUniqueNames)}\u0002${a}`)));
}

// Columns drilldown (C4, 2026-10-14): the Columns positions before the slot that are REAL fields —
// Σ Values can sit there too, but it isn't part of an expansion's outer context (the backend's rule).
function columnsOuterIndexes(columns, drilldown) {
  const slot = columns.indexOf(drilldown.topLevelField);
  return Array.from({ length: Math.max(0, slot) }, (_, i) => i).filter((i) => columns[i] !== MEASURES_PLACEHOLDER);
}

// The Columns drilldown as sent to the API: visible expansions only, scoped per outer item when real
// fields sit before the slot (stale entries from an earlier field order are dropped).
function cleanColumnsDrilldown(request) {
  const d = request.columnsHierarchyDrilldown;
  if (!d) return null;
  const outerCount = columnsOuterIndexes(request.columns, d).length;
  if (outerCount > 0) {
    return {
      topLevelField: d.topLevelField,
      expandedMemberUniqueNames: [],
      scopedExpansions: visibleScopedExpansions(d)
        .filter((e) => e.memberUniqueName.endsWith("]") && e.outerUniqueNames.length === outerCount)
        .map((e) => ({ outerUniqueNames: e.outerUniqueNames, memberUniqueName: e.memberUniqueName })),
    };
  }
  return { topLevelField: d.topLevelField, expandedMemberUniqueNames: visibleDrilldownExpansions(d).filter((u) => u.endsWith("]")) };
}

// Toggles one member of a drilldown (either axis) in `current`: per outer item when `outerUniqueNames`
// is non-empty, else in the context-free list. Like Excel, collapsing a member REMEMBERS expansions
// under it — only this member leaves the list; `ancestors` is recorded so the hidden ones are skipped.
function toggleDrilldownExpansion(current, uniqueName, outerUniqueNames, ancestors) {
  if (outerUniqueNames.length > 0) {
    const scoped = current.scopedExpansions ?? [];
    const isThis = (e) => outerKey(e.outerUniqueNames) === outerKey(outerUniqueNames) && e.memberUniqueName === uniqueName;
    const nextScoped = scoped.some(isThis) ? scoped.filter((e) => !isThis(e)) : [...scoped, { outerUniqueNames, memberUniqueName: uniqueName, ancestors }];
    return { ...current, scopedExpansions: nextScoped };
  }
  const collapsing = current.expandedMemberUniqueNames.includes(uniqueName);
  const expandedAncestors = { ...(current.expandedAncestors ?? {}) };
  if (!collapsing) expandedAncestors[uniqueName] = ancestors;
  const expanded = collapsing ? current.expandedMemberUniqueNames.filter((u) => u !== uniqueName) : [...current.expandedMemberUniqueNames, uniqueName];
  return { ...current, expandedMemberUniqueNames: expanded, expandedAncestors };
}

function isDrilldownExpansionOn(current, uniqueName, outerUniqueNames) {
  if (!current) return false;
  if (outerUniqueNames.length > 0) {
    return (current.scopedExpansions ?? []).some((e) => outerKey(e.outerUniqueNames) === outerKey(outerUniqueNames) && e.memberUniqueName === uniqueName);
  }
  return current.expandedMemberUniqueNames.includes(uniqueName);
}

// A request saved before Drilldown v3 kept a drilldown OUTSIDE Rows (Rows = []); it's a slot in Rows
// now — put it there so the rest of the app (reconcileDrilldown, the Rows pills) sees it.
function migrateDrilldownSlot(request) {
  const drilldown = request.rowsHierarchyDrilldown;
  if (!drilldown || request.rows.includes(drilldown.topLevelField)) return request;
  return { ...request, rows: [...request.rows, drilldown.topLevelField] };
}

const EMPTY_REQUEST = {
  rows: [],
  columns: [],
  values: [],
  filters: [],
  sort: null,
  rowLevelSorts: [],
  columnLevelSorts: [],
  // Real Excel's own Design tab "Grand Totals" is a 4-way dropdown (Off / On for Rows and Columns /
  // On for Rows Only / On for Columns Only), not a single on/off toggle — matches the backend's own
  // GrandTotalsOption string values exactly (added 2026-10-06, "عايز كل حاجة زي ال excel", after a
  // plain checkbox was found not to be real parity). Defaults to RowsAndColumns, matching real
  // Excel's own default the instant any PivotTable is created (confirmed live, 2026-10-06, same
  // default-behaviors audit that also caught the first-column sort gap below).
  grandTotals: "RowsAndColumns",
  // Real Excel's own Design tab "Subtotals" is a 3-way choice (Do Not Show / Show all Subtotals at
  // Bottom of Group / Show all Subtotals at Top of Group) — same gap as grandTotals above, same fix.
  // Defaults to Off, preserving this app's existing behavior (no live-Excel confirmation yet of
  // what Excel's own default position actually is — see EXCEL_PARITY_AUDIT.md).
  subtotals: "Off",
  // Real Excel's own per-field "Field Settings…" → Subtotals (2026-10-13) — overrides `subtotals`
  // above for ONE specific Rows field, falling back to it when a field has no entry here. Each entry
  // is `{ field, position }` (position: "Off"/"Bottom"/"Top", matching `subtotals`' own string
  // values exactly so both share the same <select> options). Same "list of per-field records" shape
  // `values`/`filters` already use, not a dictionary.
  fieldSubtotals: [],
  calculatedFields: [],
  calculatedItems: [],
  dateGroupings: [],
  numericGroupings: [],
  showItemsWithNoData: [],
  // Real Excel's per-field "Show items with no data" (PLAN 7, 2026-10-13, EXPERIMENTAL — see the
  // plan file's own WHY comment on why this needs live-cube verification). Unlike showItemsWithNoData
  // above (per-AXIS), this is a plain list of field names — a field in this list keeps its own
  // zero-data members even when the rest of its axis still filters normally. Query-affecting (unlike
  // fieldLayouts' Compact/RepeatItemLabels/CustomName/InsertBlankLine), so it stays inside `request`
  // and IS sent to queryPivot, not stripped out. Rows-only; v1 supports exactly one field at a time
  // (MdxPivotQueryBuilder rejects more — see its own WHY comment).
  fieldsShowingItemsWithNoData: [],
  // Real Excel's per-member OLAP drilldown (PLAN 8, 2026-10-13, EXPERIMENTAL) — { topLevelField,
  // expandedMemberUniqueNames } or null. Mutually exclusive with `rows` having any real field (see
  // FieldPicker.jsx's own startDrilldown/Rows-onChange WHY comments) — query-affecting, stays in
  // `request` and IS sent to queryPivot, same as fieldsShowingItemsWithNoData.
  rowsHierarchyDrilldown: null,
  // Columns drilldown (C4, 2026-10-14) — same shape, a slot among `columns`. Started by moving a
  // drilldown pill from Rows to Columns (see dragDrop.js reconcileDrilldowns).
  columnsHierarchyDrilldown: null,
  // Excel's own Slicer control (Insert > Slicer) — a separate button-grid widget bound to a field,
  // distinct from the Filters drop zone/Report Filter rows even though it reads/writes the SAME
  // request.filters entry (Members mode). This array is purely "which fields get the Slicer widget
  // instead of (not in addition to) a Report Filter row" — a display concern the backend has no
  // reason to know about, so it's stripped out of the cleaned request before queryPivot (see
  // runQuery below), same treatment reportLayout/tableStyle/includeChart already get.
  slicerFields: [],
  // Excel's own Timeline control (Insert > Timeline) — same "separate widget over the same
  // request.filters entry" relationship as slicerFields above, just date-range-shaped instead of a
  // button grid. Also stripped from the cleaned request before queryPivot.
  timelineFields: [],
};

// Real Excel keeps whatever you were working on even if you haven't saved it yet — switching
// sheets (or here, navigating to Snapshot History and back) doesn't blank out an in-progress
// PivotTable. This app previously had no equivalent: navigating away unmounted this whole page,
// and coming back always started from EMPTY_REQUEST, silently discarding an unsaved pivot
// (caught live 2026-10-07, "اكيد عايزه زي ال excel" after asking exactly what happens in that
// sequence). Only the lightweight CONFIGURATION is persisted (request + display settings) — never
// `result` itself (can be arbitrarily large, no data-size cap in this app — see CubeOptions' own
// WHY comment — so stashing it in localStorage risked silently hitting its ~5-10MB quota). The
// existing auto-run effect below already re-queries the cube the instant `request` is non-empty,
// so restoring just the config is enough to naturally repopulate `result` on its own.
// Exported so SnapshotHistory.jsx's "Edit" button can write a saved snapshot's own config into the
// SAME slot this page restores from on mount (see its own WHY comment above) — reusing this one
// key, rather than inventing a parallel "snapshot I'm about to edit" channel, means Edit is just
// "restore this instead of whatever was here before," no new restore code path needed.
export const BUILDER_STATE_STORAGE_KEY = "pivotSnapshot.builderState";

function readPersistedBuilderState() {
  try {
    const raw = localStorage.getItem(BUILDER_STATE_STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    // Per-browser convenience only — same no-op-on-failure as readEnabledTables above; a
    // corrupted/unreadable entry just means starting from EMPTY_REQUEST, same as before this existed.
    return null;
  }
}

// PivotResult only carries member VALUES (e.g. "Cardiology"), not the field's own name — this
// looks each row/column field id up in the dimensions list to get something readable ("Specialty
// Name") for the grid's header row, instead of leaving it blank.
function fieldLabels(fieldIds, dimensions) {
  if (!fieldIds) return [];
  return fieldIds.map((id) =>
    id === MEASURES_PLACEHOLDER ? "Σ Values" : (dimensions.find((d) => d.field === id)?.displayName ?? id),
  );
}

// Which tables the "Available fields" picker actually works with, out of everything the cube's
// schema has (302 tables as of 2026-09-27, mostly internal/system tables no business user needs).
// Kept in the browser, not the backend: the user asked for a per-user pick ("خلي اليوزر يختار"),
// not one shared setting for everyone.
const ENABLED_TABLES_STORAGE_KEY = "pivotSnapshot.enabledTables";

function readEnabledTables() {
  try {
    const raw = localStorage.getItem(ENABLED_TABLES_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

// A filter is only ready to send once whichever fields ITS mode actually needs are filled in —
// each mode uses a different subset of PivotFilter's optional fields (see the backend's
// MdxPivotQueryBuilder.ValidateFilters for the authoritative per-mode rules this mirrors).
function isFilterComplete(filter) {
  switch (filter.mode ?? "Members") {
    case "LabelContains":
    case "LabelBeginsWith":
    case "LabelEndsWith":
      return Boolean(filter.labelText);
    case "TopN":
    case "BottomN":
      return Boolean(filter.n > 0 && filter.byMeasureField);
    default:
      return filter.includedMembers.length > 0;
  }
}

// Everything in `request` EXCEPT each value's `format`/`conditionalFormat` — those two are pure
// display settings (see displayValues' own WHY comment below) that must NOT mark the on-screen
// result as stale, unlike every other field here (rows, columns, filters, sort, groupings,
// totals, showItemsWithNoData, calculatedFields/Items, and even a value's
// showValuesAs — that one DOES need a fresh query since it's computed server-side). Used to detect
// "the user changed something that actually invalidates the table currently on screen."
function structuralSignature(req) {
  return JSON.stringify({
    ...req,
    values: req.values.map(({ format: _format, conditionalFormat: _conditionalFormat, ...structural }) => structural),
  });
}

export default function PivotBuilder() {
  // Merged with EMPTY_REQUEST's own defaults (not a bare `?? EMPTY_REQUEST` fallback) — a request
  // blob saved to localStorage BEFORE a newer field existed (fieldSubtotals, fieldLayouts,
  // fieldsShowingItemsWithNoData, rowsHierarchyDrilldown, …) would otherwise restore with that key
  // simply MISSING rather than defaulted, and every `request.<newField>.filter(...)`-style call
  // throughout this file assumes it's always at least an empty array/null, not undefined — caught
  // live, 2026-10-13, the first time a stale pre-PLAN-7 localStorage blob hit one of those call
  // sites ("Cannot read properties of undefined (reading 'filter')"). A shallow merge is enough:
  // every one of these fields is a plain top-level array/scalar/null, never itself a nested object
  // whose OWN missing sub-keys would need defaulting too.
  const [request, setRequestRaw] = useState(() => migrateDrilldownSlot({ ...EMPTY_REQUEST, ...readPersistedBuilderState()?.request }));
  // Wraps every setRequest call with normalizeMeasuresPlacement so `request` itself ALWAYS has Σ
  // Values at an explicit position (defaulting to the end of Columns) the instant at least one
  // Value field exists — see that function's own WHY comment for the real bug this fixes (silently
  // broken measureIndexOf-based features — the value-cell context menu, Number Format, Conditional
  // Formatting, axis Sort-by-measure — for any pivot where the user never manually dragged the Σ
  // Values pseudo-field themselves, by far the common case). Wrapping the SETTER (not calling
  // normalizeMeasuresPlacement at every read site) fixes it at the one source every caller already
  // goes through, instead of auditing every FieldPicker/PivotToolbar/PivotGrid prop individually.
  const setRequest = (updater) => {
    setRequestRaw((prev) => {
      const next = normalizeMeasuresPlacement(typeof updater === "function" ? updater(prev) : updater);
      // Real Excel's own default the moment an axis gets its first field — every PivotTable row AND
      // column area starts sorted Ascending (A→Z / chronological) by label, never "whatever order
      // the cube happens to return" (confirmed live, 2026-10-06: "انا اتقالي انه في ال excel دايما
      // بيرتب اول column", then "لا عايز نفس ال excel" when a Rows-only version was floated as a
      // compromise over the Columns side — both axes, not just Rows). Routed through
      // RowLevelSorts/ColumnLevelSorts (Level 0, no ByMeasureField = sort by label) rather than the
      // single shared `sort` field: `sort` can only target ONE axis at a time and is mutually
      // exclusive, GLOBALLY, with any non-empty *LevelSorts (MdxPivotQueryBuilder.ValidateLevelSorts)
      // — but RowLevelSorts and ColumnLevelSorts freely coexist with EACH OTHER, so that's the only
      // mechanism that can default both axes independently at once. Each half only fires on that
      // axis's own genuine empty→non-empty transition, and only when nothing already claims sorting
      // for the WHOLE request (`sort` still null) and that axis's own LevelSorts are still empty — an
      // explicit "Clear sort"/LevelSortEditor rule, on either axis, is left untouched.
      // Per axis since 2026-10-15: a `sort` only blocks the default on its OWN axis (the backend now
      // lets a Sort and LevelSorts coexist on different axes).
      let withDefaultSort = next;
      {
        const rowsJustGotFirstField = prev.rows.length === 0 && withDefaultSort.rows.length > 0;
        // Not for a drilldown slot (Drilldown v3: starting a drilldown on empty Rows makes it the
        // first Rows field) — a per-level sort can't apply to a drilldown axis (the backend rejects
        // it); the hierarchy keeps its own natural order unless sorted from its header.
        if (rowsJustGotFirstField && withDefaultSort.sort?.axis !== "Rows" && withDefaultSort.rowLevelSorts.length === 0 && !withDefaultSort.rowsHierarchyDrilldown) {
          withDefaultSort = { ...withDefaultSort, rowLevelSorts: [{ level: 0, direction: "Ascending", byMeasureField: null }] };
        }
        const prevColumnFieldCount = prev.columns.filter((f) => f !== MEASURES_PLACEHOLDER).length;
        const nextColumnFieldCount = withDefaultSort.columns.filter((f) => f !== MEASURES_PLACEHOLDER).length;
        const columnsJustGotFirstField = prevColumnFieldCount === 0 && nextColumnFieldCount > 0;
        // Same exception as Rows for a Columns drilldown slot (C4).
        if (columnsJustGotFirstField && withDefaultSort.sort?.axis !== "Columns" && withDefaultSort.columnLevelSorts.length === 0 && !withDefaultSort.columnsHierarchyDrilldown) {
          withDefaultSort = { ...withDefaultSort, columnLevelSorts: [{ level: 0, direction: "Ascending", byMeasureField: null }] };
        }
      }
      // A manually-set label/custom-order Sort (byMeasureField AND customOrder both null) is only
      // valid when its OWN axis has exactly one real field — the backend's ValidateSort rejects it
      // outright otherwise ("Sorting by label requires exactly one row/column field..."), since it
      // can't tell which nesting level the label refers to. SortEditor's own UI already can't produce
      // this (it only offers a single-field axis-wide label sort when that axis genuinely has one
      // field), but adding a SECOND field to an axis already carrying one is a very normal next step
      // that would otherwise leave a now-invalid `sort` sitting in the request until Run Query throws
      // — caught in the same 2026-10-06 audit that replaced the Rows-only version of the default
      // above with this one. Clearing it here (rather than migrating it to RowLevelSorts/
      // ColumnLevelSorts) matches how every other sort-invalidating edit in this app already behaves
      // — FieldPicker's own Rows/Columns onChange already clears rowLevelSorts/columnLevelSorts
      // outright on any add/remove/reorder, never tries to migrate them forward.
      if (withDefaultSort.sort && withDefaultSort.sort.byMeasureField == null && withDefaultSort.sort.customOrder == null) {
        const axisFields = withDefaultSort.sort.axis === "Rows" ? withDefaultSort.rows : withDefaultSort.columns;
        const realFieldCount = axisFields.filter((f) => f !== MEASURES_PLACEHOLDER).length;
        // With a drilldown, a Rows label sort always targets the HIERARCHY (siblings within each
        // parent — mirrors MdxPivotQueryBuilder.ValidateSort, which counts the drilldown as the one
        // field), however many other Rows fields sit beside it. Without this, sorting "Row Labels"
        // A→Z/Z→A was silently cleared right here (caught live 2026-10-14).
        const targetsDrilldown =
          (withDefaultSort.sort.axis === "Rows" && withDefaultSort.rowsHierarchyDrilldown) ||
          (withDefaultSort.sort.axis === "Columns" && withDefaultSort.columnsHierarchyDrilldown);
        if (realFieldCount !== 1 && !targetsDrilldown) {
          return { ...withDefaultSort, sort: null };
        }
      }
      return withDefaultSort;
    });
  };
  const [lastQuery, setLastQuery] = useState(null); // the exact cleaned request that produced `result`
  const [result, setResult] = useState(null);
  // structuralSignature(request) as of the last successful run — compared on every `request` change
  // below to clear a stale result instead of leaving a table on screen that no longer matches the
  // current Rows/Columns/Filters/etc (explicitly requested, 2026-09-27: a table left over from a
  // previous run was confusing once the fields it was built from had already changed).
  const [lastRunSignature, setLastRunSignature] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [measures, setMeasures] = useState([]);
  const [metaError, setMetaError] = useState(null);
  // Separate from `loading` above (that one's for Run query). `tables` is the fast, columns-free
  // catalog (302 tables) the multiselect below picks from; `fieldsByTable` is a cache of each
  // table's own columns ever fetched, fetched lazily one table at a time instead of eagerly
  // loading every table's columns up front (2301 fields, several seconds — explicitly asked for
  // live, 2026-09-27, after the earlier "load everything" version made the panel sit empty that
  // whole time with nothing saying so).
  const [tables, setTables] = useState([]);
  const [tablesLoading, setTablesLoading] = useState(true);
  const [enabledTableNames, setEnabledTableNames] = useState(readEnabledTables);
  const [fieldsByTable, setFieldsByTable] = useState({});
  const [fieldsLoading, setFieldsLoading] = useState(false);
  // Filtered down to CURRENTLY enabled tables, not the whole cache — un-picking a table in the
  // multiselect must make it (and its columns) disappear from the picker right away, even though
  // fieldsByTable itself keeps its columns cached so re-picking it later is instant instead of a
  // fresh fetch (caught live, 2026-09-27: un-picking "Diagnosis" left it sitting in "Selected
  // tables" below since dimensions was built from the whole cache instead of just what's enabled).
  const enabledDimensions = enabledTableNames.flatMap((name) => fieldsByTable[name] ?? []);
  // `dimensions` is used for FAR more than the Available Fields tree's own display — AxisGroupingEditor's
  // date/numeric filtering, FilterRow's field dropdown, every zone's isFieldAccepted check, and
  // the Values auto-Sum-vs-Count default all key off it too. "Collapse All" (or collapsing one
  // table) calls onEnabledTableNamesChange([]) / removes just that table, which used to make EVERY
  // already-placed field from that table silently vanish from `dimensions` — breaking its grouping
  // control, its filter dropdown display, and cross-zone drag for it, even though the field was
  // still very much live in the pivot and its data was still sitting right there in fieldsByTable's
  // own cache (caught 2026-10-06, full codebase review — the "Collapse All" comment's own claim that
  // this "doesn't touch which fields are actually placed on the pivot" turned out to be wrong in
  // practice). Fields currently placed anywhere in the request are re-added back in from the FULL
  // cache (searched across every table ever fetched, not just the enabled ones) if their own table
  // got collapsed/un-enabled out from under them.
  const enabledFieldIds = new Set(enabledDimensions.map((d) => d.field));
  const placedFieldIds = new Set([
    ...request.rows.filter((f) => f !== MEASURES_PLACEHOLDER),
    ...request.columns.filter((f) => f !== MEASURES_PLACEHOLDER),
    ...request.filters.map((f) => f.field),
    ...request.values.filter((v) => v.field.startsWith("[")).map((v) => v.field),
  ]);
  const stillPlacedButCollapsed = Object.values(fieldsByTable)
    .flat()
    .filter((d) => placedFieldIds.has(d.field) && !enabledFieldIds.has(d.field));
  const dimensions = [...enabledDimensions, ...stillPlacedButCollapsed];

  const updateEnabledTables = (names) => {
    setEnabledTableNames(names);
    try {
      localStorage.setItem(ENABLED_TABLES_STORAGE_KEY, JSON.stringify(names));
    } catch {
      // Per-browser convenience only — nothing breaks if this can't be saved (private window,
      // storage disabled, quota); the picker just won't remember the choice next time.
    }
  };

  // Number format and conditional format are pure display settings — computed entirely client-side
  // from cell values already in `result`, no cube round-trip needed. But the grid/chart/export/save
  // below all key off `lastQuery.values` (the frozen snapshot from the last successful query, whose
  // field order/count actually matches `result`'s columns) rather than the live `request.values` —
  // so picking Color scale in a value pill's popover looked like it silently did nothing until you
  // clicked Run query again (caught live, 2026-09-27). Overlaying just `format`/`conditionalFormat`
  // from the live request onto lastQuery's values (matched by field name) fixes that without risking
  // a mismatch when the live request has ADDED/REMOVED a value field entirely — that still correctly
  // needs a fresh Run query, since `result` wouldn't have that field's column yet.
  const displayValues = lastQuery?.values?.map((v) => {
    const live = request.values.find((rv) => rv.field === v.field);
    return live ? { ...v, format: live.format, conditionalFormat: live.conditionalFormat } : v;
  });

  const [snapshotName, setSnapshotName] = useState("");
  // Optional label for just THIS version (e.g. "Q3 Final") — shown alongside, not instead of, the
  // auto-incrementing Version N (requested live 2026-10-07, "دا او دا" in answer to "replace the
  // number or add next to it" — read as "support both": a blank name still shows plain "Version N",
  // exactly as before this existed, so it's purely additive).
  const [versionName, setVersionName] = useState("");
  // Version name input starts hidden behind a small "+ Add..." button, only offered once a
  // Snapshot name has actually been typed (requested live 2026-10-12: "خلي ال snapshot name بس
  // وجواه... لو عايز version name ولو لا يبقى خلاص") — two always-visible stacked inputs for what's
  // usually a one-field decision (most saves don't need a version label at all) was more UI than the
  // common case needed, now that this sidebar is otherwise a tight, vertically-stacked column.
  const [showVersionName, setShowVersionName] = useState(false);
  // Existing snapshot names, offered as <datalist> suggestions on the name field below — native
  // HTML combo behavior (pick one of these OR type anything else) needs no custom dropdown
  // component, and typing a name that doesn't match any of these is exactly how a NEW snapshot is
  // created today anyway (requested live 2026-10-07: "لو اختار الـ snapshot name اللي موجودة قبل
  // كدا او لو مش موجودة تبقى new"). Fetched once on mount — good enough for a list that only
  // changes when someone saves a new one, not worth re-fetching on every keystroke/save.
  const [existingSnapshotNames, setExistingSnapshotNames] = useState([]);
  useEffect(() => {
    listDefinitions()
      .then((defs) => setExistingSnapshotNames(defs.map((d) => d.name)))
      .catch(() => {
        // Non-critical — the name field still works as a plain free-text input without
        // suggestions, same as before this existed.
      });
  }, []);
  const [tableStyle, setTableStyle] = useState(() => readPersistedBuilderState()?.tableStyle ?? "Default");
  const [includeChart, setIncludeChart] = useState(() => readPersistedBuilderState()?.includeChart ?? false);
  const [chartType, setChartType] = useState(() => readPersistedBuilderState()?.chartType ?? "Bar");
  // Real Excel's own default for a brand-new PivotTable is Compact Form, not Tabular — same
  // 2026-10-06 default-behaviors audit as the Grand Totals default above.
  const [reportLayout, setReportLayout] = useState(() => readPersistedBuilderState()?.reportLayout ?? "Compact");
  // "Repeat All Item Labels" (PivotTable Design tab) — off by default, same as real Excel. Only
  // meaningful for Outline/Tabular (Compact has no separate per-level column to repeat a label
  // into, same as real Excel graying the option out there — see PivotGrid.jsx's own WHY comment).
  const [repeatItemLabels, setRepeatItemLabels] = useState(() => readPersistedBuilderState()?.repeatItemLabels ?? false);
  // Real Excel's per-field "Field Settings…" > Custom Name / Layout & Print (PLAN 6 Phase 2,
  // 2026-10-13) — one entry per Rows field with its own Compact/RepeatItemLabels/InsertBlankLine/
  // CustomName override. Deliberately a sibling of reportLayout/repeatItemLabels above, NOT part of
  // `request` — none of these four change which rows/cells the backend returns, only how they're
  // displayed, so (like reportLayout/repeatItemLabels) they must stay out of structuralSignature or
  // every per-field display toggle would needlessly re-query the live cube.
  const [fieldLayouts, setFieldLayouts] = useState(() => readPersistedBuilderState()?.fieldLayouts ?? []);
  // Excel's value-cell context menu has "Hide Field List" — toggles the FIELDS sidebar's own
  // visibility, with a small floating button to bring it back (Excel keeps the same toggle on the
  // PivotTable Analyze ribbon too, not just the context menu, but there's no ribbon button for it
  // here yet — this is the only entry point so far).
  const [showFieldList, setShowFieldList] = useState(true);
  const [saving, setSaving] = useState(false);
  const [saveMessage, setSaveMessage] = useState(null);
  const [saveError, setSaveError] = useState(null);
  const [exporting, setExporting] = useState(false);
  const [exportError, setExportError] = useState(null);

  useEffect(() => {
    Promise.all([getTables(), getMeasures()])
      .then(([tbls, meas]) => {
        setTables(tbls);
        setMeasures(meas);
      })
      .catch((err) => {
        // A rejection here isn't always a plain Error with a real .message (caught live,
        // 2026-09-27: a concurrent-token-refresh race — see dataverseAuth.js's WHY comment —
        // produced one whose falsy .message made {metaError && (...)} render nothing at all,
        // a completely silent failure with no visible error anywhere). Falling back through
        // toString()/JSON.stringify guarantees something is always shown here.
        setMetaError(String(err?.message || err?.toString?.() || JSON.stringify(err) || "Unknown error."));
      })
      .finally(() => setTablesLoading(false));
  }, []);

  // Lazily fetches columns for any ENABLED table not already in the cache — fires once on mount
  // for whatever was already picked in a previous session (from localStorage) and again each time
  // the user adds a table in the multiselect. Never re-fetches a table already cached, and never
  // drops a table's columns from the cache just because it was un-picked (re-picking it later is
  // then instant, and any field from it still placed on the pivot keeps working).
  //
  // enabledTableNames holds DISPLAY names ("Doctor"), matching what the multiselect shows and what
  // FieldPicker's own column-grouping already keys on — but getDimensionsForTable needs the cube's
  // raw identifier ("[Dim Doctor]"), only known once `tables` itself has loaded. Waiting on
  // `tables.length` below avoids caching an empty result for a name whose real table just hasn't
  // arrived yet (that name would otherwise look "already fetched" and never be retried).
  useEffect(() => {
    if (tables.length === 0) return;

    const missing = enabledTableNames.filter((name) => !(name in fieldsByTable));
    if (missing.length === 0) return;

    setFieldsLoading(true);
    Promise.all(
      missing.map((name) => {
        const table = tables.find((t) => t.displayName === name);
        return table ? getDimensionsForTable(table.field).then((fields) => [name, fields]) : Promise.resolve([name, []]);
      }),
    )
      .then((entries) => {
        setFieldsByTable((prev) => {
          const next = { ...prev };
          for (const [name, fields] of entries) next[name] = fields;
          return next;
        });
      })
      .catch((err) => {
        setMetaError(String(err?.message || err?.toString?.() || JSON.stringify(err) || "Unknown error."));
      })
      .finally(() => setFieldsLoading(false));
  }, [enabledTableNames, fieldsByTable, tables]);

  // Drop the table currently on screen the moment the user changes anything that would actually
  // change the query (as opposed to a pure display tweak like Format/Conditional format, which
  // structuralSignature ignores) — otherwise it just sits there looking current while it's really
  // built from Rows/Columns/Filters/etc that no longer match what's selected (explicitly requested,
  // 2026-09-27). Save/export messages are cleared alongside it since they describe that now-gone
  // result too.
  useEffect(() => {
    if (result && structuralSignature(request) !== lastRunSignature) {
      setResult(null);
      setLastQuery(null);
      setSaveMessage(null);
      setExportError(null);
    }
  }, [request, result, lastRunSignature]);

  // Auto-run (see the effect below) can fire a new query while a previous one is still in flight
  // — e.g. the user drops a second field before the first query's response lands. Whichever
  // response arrives LAST would otherwise win regardless of which request is actually newer,
  // intermittently showing a stale result. Each call claims the next token and only commits state
  // if it's still the latest when its response lands; an outrun call's response is silently
  // dropped (the newer call already owns loading/error/result).
  const queryTokenRef = useRef(0);
  // The queryTokenRef pattern above already stops a STALE response from being applied to state —
  // but it does nothing to stop the ABANDONED request itself from continuing to run on the shared
  // cube server after nobody's waiting on it anymore. This aborts the actual in-flight fetch for
  // the previous call the moment a newer one starts, so the backend's own CubeQueryService (which
  // now calls AdomdCommand.Cancel() when its request is aborted — see its own WHY comment) stops
  // doing that wasted work instead of finishing it regardless (requested live, 2026-10-05, "عايز
  // تزود cancellation token").
  const queryAbortControllerRef = useRef(null);
  // Also abort whatever's in flight if the page itself is torn down mid-query (navigating away,
  // hot reload, etc.) — same reasoning as the per-call abort in runQuery below, just for the
  // "nobody's calling runQuery again at all" case that abort wouldn't otherwise cover.
  useEffect(() => () => queryAbortControllerRef.current?.abort(), []);

  const runQuery = async () => {
    const token = ++queryTokenRef.current;
    // Whatever the PREVIOUS runQuery call is still waiting on is now definitely stale (this call's
    // own token just superseded it) — abort its request outright instead of letting it keep
    // running on the cube only to be ignored by the queryTokenRef check below once it does land.
    queryAbortControllerRef.current?.abort();
    const controller = new AbortController();
    queryAbortControllerRef.current = controller;
    setLoading(true);
    setError(null);
    setSaveMessage(null);
    try {
      const cleaned = {
        rows: request.rows.filter(Boolean),
        columns: request.columns.filter(Boolean),
        values: request.values.filter((v) => v.field),
        filters: request.filters.filter((f) => f.field && isFilterComplete(f)),
        sort: request.sort,
        rowLevelSorts: request.rowLevelSorts,
        columnLevelSorts: request.columnLevelSorts,
        grandTotals: request.grandTotals,
        subtotals: request.subtotals,
        // Defensive, same reasoning as every other per-field array cleaned here — a field's own
        // override is meaningless (and the backend's own FieldSubtotals model doesn't expect it)
        // once that field has actually left Rows.
        // A null position means "follow the report-wide Subtotals position" (a Custom field keeps
        // its functions but not its own position) — resolved here because the backend's
        // FieldSubtotals.Position is required (null was a 400, caught live 2026-10-14).
        // Also kept for a Columns drilldown's own field (its "Egypt Total" columns, 2026-10-14); never
        // for a Rows drilldown slot (the backend rejects an override there).
        fieldSubtotals: request.fieldSubtotals
          .filter((f) =>
            request.columns.includes(f.field) ||
            (request.rows.includes(f.field) && request.rowsHierarchyDrilldown?.topLevelField !== f.field))
          .map((f) => ({ ...f, position: f.position ?? request.subtotals })),
        calculatedFields: request.calculatedFields.filter(isCompleteCalculatedField),
        // A CalculatedItem needs a field, a name, and at least one member on either side (see
        // MdxPivotQueryBuilder.ValidateCalculatedItems) — mirrors the calculatedFields cleaning
        // above.
        calculatedItems: request.calculatedItems.filter(isCompleteCalculatedItem),
        // A grouped field can be in Rows OR Columns (see GroupingProcessor) — checking only Rows
        // here would silently strip a Columns-axis grouping right before it's sent.
        dateGroupings: request.dateGroupings.filter((g) => request.rows.includes(g.field) || request.columns.includes(g.field)),
        numericGroupings: request.numericGroupings.filter(
          (g) => (request.rows.includes(g.field) || request.columns.includes(g.field)) && g.binSize > 0,
        ),
        showItemsWithNoData: request.showItemsWithNoData,
        fieldsShowingItemsWithNoData: request.fieldsShowingItemsWithNoData.filter((f) => request.rows.includes(f)),
        // Defensive, same reasoning as every other array cleaned here — but for a real data-quality
        // issue, not a stale-reference one: some members on this cube return an incomplete
        // UniqueName (ending in a bare "&" with no key after it). PivotGrid.jsx already refuses to
        // render a +/- toggle for one of these, but an ALREADY-expanded one from before that guard
        // existed can still be sitting in localStorage-persisted state — caught live, 2026-10-13,
        // "عملت Refresh ولسه نفس المشكلة": the stale entry survives a refresh since it's part of the
        // persisted request, not component state. Filtering it out here means it self-heals on the
        // very next query instead of needing a manual Reset every time this specific cube quirk
        // happens to recur.
        //
        // Only VISIBLE expansions are sent (Drilldown v2, 2026-10-14): Excel remembers a member's
        // expansion while one of its ancestors is collapsed, so state keeps every expanded member
        // (plus its ancestry, frontend-only) and the query gets just those whose ancestors are all
        // expanded — otherwise the hidden member's children would show up orphaned under the
        // collapsed ancestor.
        //
        // Drilldown v3: with Rows fields BEFORE the slot, expansions are per outer item and go as
        // scopedExpansions (the backend requires exactly that split — see ValidateCombinedDrilldown).
        rowsHierarchyDrilldown:
          request.rowsHierarchyDrilldown &&
          (request.rows.indexOf(request.rowsHierarchyDrilldown.topLevelField) > 0
            ? {
                topLevelField: request.rowsHierarchyDrilldown.topLevelField,
                expandedMemberUniqueNames: [],
                scopedExpansions: visibleScopedExpansions(request.rowsHierarchyDrilldown)
                  .filter((e) => e.memberUniqueName.endsWith("]"))
                  .map((e) => ({ outerUniqueNames: e.outerUniqueNames, memberUniqueName: e.memberUniqueName })),
              }
            : {
                topLevelField: request.rowsHierarchyDrilldown.topLevelField,
                expandedMemberUniqueNames: visibleDrilldownExpansions(request.rowsHierarchyDrilldown).filter((u) => u.endsWith("]")),
              }),
        // Columns drilldown (C4): the same visible-only rule; scoped exactly when REAL fields (not
        // Σ Values) sit before its slot — mirrors MdxPivotQueryBuilder.ValidateColumnsDrilldown.
        columnsHierarchyDrilldown: cleanColumnsDrilldown(request),
      };
      const data = await queryPivot(cleaned, { signal: controller.signal });
      if (token !== queryTokenRef.current) return; // a newer call already landed/owns the UI
      setResult(data);
      setLastQuery(cleaned);
      setLastRunSignature(structuralSignature(request));
    } catch (err) {
      if (token !== queryTokenRef.current) return;
      setError(err.message);
      setResult(null);
      setLastQuery(null);
    } finally {
      if (token === queryTokenRef.current) setLoading(false);
    }
  };

  // Always the latest runQuery, read from inside the effect below without making the effect
  // depend on it — runQuery is a plain closure re-created every render (not useCallback'd), so
  // depending on it directly would restart the debounce timer on every render for reasons
  // unrelated to the request actually changing (loading flipping, ...), which would fight the
  // whole point of debouncing. Assigned in its own effect (no deps array = runs after every
  // render), not during render itself — React's own refs rule.
  const runQueryRef = useRef(runQuery);
  useEffect(() => {
    runQueryRef.current = runQuery;
  });

  // Auto-run: real Excel re-queries the instant you drop a field, no explicit "Run" step — this
  // app talks to a live SSAS server instead of local data, so a query actually costs something,
  // which is why a debounce exists here at all (requested live,
  // 2026-10-04: "تشغيل تلقائي عند كل سحب/إفلات", after "فين زرار Run query في الإكسل؟"). Keyed off
  // structuralSignature, not `request` itself, so a pure Format/Conditional format edit never
  // triggers a query — same reasoning, same signature, as the stale-result-clearing effect above.
  // Skipped while the pivot is still empty (nothing in Rows/Columns/Values) — Excel shows nothing
  // until then either, no point querying for it.
  const requestSignature = structuralSignature(request);
  // Real Excel's per-member OLAP drilldown (PLAN 8) leaves request.rows EMPTY on purpose
  // (RowsHierarchyDrilldown replaces it entirely) — without this check, clicking "Drilldown" set
  // rowsHierarchyDrilldown but rows/columns/values all stayed empty, so isPivotEmpty stayed true and
  // the auto-run effect below silently never queried at all (caught live, 2026-10-13, "بدوس عليها
  // مش بيحصل حاجة" — the state WAS changing, the query just never fired).
  const isPivotEmpty =
    request.rows.length === 0 && request.columns.length === 0 && request.values.length === 0 && !request.rowsHierarchyDrilldown;
  useEffect(() => {
    if (isPivotEmpty) return undefined;
    const timer = setTimeout(() => {
      runQueryRef.current();
    }, 700);
    return () => clearTimeout(timer);
  }, [requestSignature, isPivotEmpty]);

  // Persists the live config (not `result` — see BUILDER_STATE_STORAGE_KEY's own WHY comment)
  // so navigating away and back (or a page refresh) restores it instead of starting blank. Keyed
  // off the full `request` (not structuralSignature, which deliberately drops format/
  // conditionalFormat) — those two ARE worth remembering too, unlike for the auto-run effect above,
  // which only cares about changes that actually invalidate the on-screen result.
  useEffect(() => {
    try {
      localStorage.setItem(
        BUILDER_STATE_STORAGE_KEY,
        JSON.stringify({ request, tableStyle, includeChart, chartType, reportLayout, repeatItemLabels, fieldLayouts }),
      );
    } catch {
      // Per-browser convenience only — same no-op-on-failure as every other localStorage write in
      // this file; worst case, the next visit just starts blank again.
    }
  }, [request, tableStyle, includeChart, chartType, reportLayout, repeatItemLabels, fieldLayouts]);

  // Orphaned-entry cleanup for fieldLayouts, same reasoning as request.fieldSubtotals' own cleanup
  // in FieldPicker.jsx (moveField/addHierarchyToAxis) — a field that leaves Rows should take its
  // Layout override with it. Done here instead of inside FieldPicker's onChange handlers because
  // fieldLayouts lives OUTSIDE `request` (see its own WHY comment above), so FieldPicker's
  // request-only onChange has no way to touch it directly.
  useEffect(() => {
    setFieldLayouts((prev) => prev.filter((f) => request.rows.includes(f.field)));
  }, [request.rows]);

  // Clears the whole builder back to a blank slate — fields, result, every warning/message, and
  // the save-panel inputs. Doesn't touch dimensions/measures (still the same cube, no need to
  // re-fetch metadata).
  //
  // Previously also called setCardinalityWarning/setCellsWarning/setConfirmedOverrides — leftover
  // calls to state that had already been removed elsewhere (the "remove every data-fetching limit"
  // decision, see CubeOptions' own WHY comment), left orphaned here with nothing declaring them
  // anymore. Clicking Reset threw "ReferenceError: setCardinalityWarning is not defined" and
  // crashed to the ErrorBoundary fallback every single time (caught live 2026-10-11 via the dev
  // server's own console output, not a user report — this had been broken since whenever that
  // cleanup happened).
  //
  // DOES now abort any in-flight query and clear `loading` immediately — previously left both
  // alone on the theory that "a reset mid-query just lets that request finish and get its result
  // silently discarded when it lands, same as any other stale-response case" (true for the RESULT,
  // which IS correctly discarded), but that left the "Updating…" spinner visibly running next to
  // the just-reset "choose fields to get started" empty state for however long the abandoned query
  // (or, worse, a genuinely hung one against a struggling cube) took to actually settle — caught
  // live, 2026-10-06, "عملت reset وهو لسا بيعمل updating". Bumping queryTokenRef here too so that
  // request's own catch/finally (keyed on that token) can't re-show loading/an error for a query
  // Reset has already disowned, same guard every other superseding call already relies on.
  // Resetting these 6 states back to their defaults also naturally re-fires the persistence effect
  // above with those same defaults, overwriting BUILDER_STATE_STORAGE_KEY back to "blank" — no
  // separate localStorage.removeItem needed here.
  const handleReset = () => {
    queryTokenRef.current += 1;
    queryAbortControllerRef.current?.abort();
    setLoading(false);
    setRequest(EMPTY_REQUEST);
    setLastQuery(null);
    setResult(null);
    setError(null);
    setSnapshotName("");
    setVersionName("");
    setShowVersionName(false);
    setTableStyle("Default");
    setIncludeChart(false);
    setChartType("Bar");
    setReportLayout("Compact");
    setRepeatItemLabels(false);
    setSaveMessage(null);
    setSaveError(null);
    setExportError(null);
  };

  // Real Excel's own behavior: drag a field out of the Field List and drop it ANYWHERE else on the
  // screen (not just back onto the Available Fields panel specifically) and it's removed from the
  // pivot — flagged live, 2026-10-04 ("لو عملت drag من المربع لـ أي حته برا المفروض يعمل زي
  // الإكسل"). Every recognized zone's own onDrop already stops propagation (see each one's own WHY
  // comment in FieldPicker.jsx/HierarchicalCubePathSelect.jsx), so this page-level handler only
  // ever sees a drop that landed nowhere FieldPicker itself recognizes as a valid target.
  const handlePageDragOver = (e) => {
    if (e.dataTransfer.types.includes(FIELD_DRAG_MIME)) {
      e.preventDefault();
    }
  };

  const handlePageDrop = (e) => {
    const payload = readFieldDragPayload(e);
    if (!payload) return;
    e.preventDefault();
    setRequest((prev) => removeFieldEverywhere(prev, payload.field));
  };

  const handleSave = async () => {
    if (!snapshotName.trim() || !result || !lastQuery) {
      return;
    }
    setSaving(true);
    setSaveError(null);
    setSaveMessage(null);
    try {
      // The SAME bundle shape (and the SAME live `request`, not the cleaned/filtered one Run Query
      // sends) the persistence effect above already writes to localStorage — reused as-is here so
      // Edit later restores this exactly as it looked while being edited, not just a filtered/
      // query-ready subset of it. See BUILDER_STATE_STORAGE_KEY's own WHY comment.
      const builderStateJson = JSON.stringify({ request, tableStyle, includeChart, chartType, reportLayout, repeatItemLabels, fieldLayouts });
      const saved = await saveSnapshot(
        snapshotName.trim(),
        { ...lastQuery, values: displayValues },
        result,
        tableStyle,
        includeChart,
        chartType,
        reportLayout,
        repeatItemLabels,
        fieldLayouts,
        builderStateJson,
        versionName.trim() || null,
      );
      setSaveMessage(`Saved as version ${saved.versionNumber}.`);
      // A successful save means this snapshot name now definitely exists — refreshed so the NEXT
      // save's datalist offers it too, without needing a page reload (e.g. saving a brand-new name
      // for the first time).
      if (!existingSnapshotNames.includes(snapshotName.trim())) {
        setExistingSnapshotNames((prev) => [...prev, snapshotName.trim()]);
      }
    } catch (err) {
      // SnapshotTooLarge included — a flat, non-overridable ceiling now (removed 2026-09-27 on
      // request), so this just shows like any other save error rather than a confirm-and-retry.
      setSaveError(err.message);
    } finally {
      setSaving(false);
    }
  };

  const handleExport = async () => {
    if (!result || !lastQuery) return;
    // A per-member drilldown pivot (PLAN 8) exports since 2026-10-14 (Drilldown v2): exportExcel.js
    // draws it through the same display-rows step as the live grid, in `drilldown` mode.
    setExporting(true);
    setExportError(null);
    try {
      // Dynamic import: ExcelJS is a genuinely large library (bundle jumped from ~420KB to
      // ~1.35MB when statically imported, 2026-09-23) — loading it only when someone actually
      // clicks Export keeps it out of everyone else's initial page load.
      const { exportPivotToExcel } = await import("../utils/exportExcel");
      await exportPivotToExcel({
        result,
        rows: lastQuery.rows,
        columns: lastQuery.columns,
        rowFieldLabels: fieldLabels(lastQuery.rows, dimensions),
        columnFieldLabels: fieldLabels(hideMeasuresLabelIfSingleValue(lastQuery.columns, lastQuery.values.length), dimensions),
        valueFields: displayValues,
        filename: snapshotName.trim() || "Pivot",
        // Same layout as the live grid, so the export draws the same rows.
        layout: reportLayout,
        repeatItemLabels,
        fieldLayouts,
        // A drilldown alone keeps its single ragged column; combined with other Rows fields
        // (Drilldown v3) it's a normal field slot, indented by member depth.
        ...(() => {
          const d = lastQuery.rowsHierarchyDrilldown;
          if (!d) return {};
          const others = lastQuery.rows.filter((f) => f !== MEASURES_PLACEHOLDER && f !== d.topLevelField);
          return others.length === 0 ? { drilldown: true } : { drillSlot: lastQuery.rows.indexOf(d.topLevelField) };
        })(),
        drilldownSubtotals: lastQuery.subtotals,
        // Columns drilldown (C4): header rows from buildColumnHeaderRows, same as the grid.
        columnsDrillSlot: lastQuery.columnsHierarchyDrilldown ? lastQuery.columns.indexOf(lastQuery.columnsHierarchyDrilldown.topLevelField) : -1,
      });
    } catch (err) {
      setExportError(err.message);
    } finally {
      setExporting(false);
    }
  };

  return (
    // onDragOver/onDrop here, outside AppShell entirely: catches a field dropped ANYWHERE on the
    // page (including the sidebar, the grid, blank margins) that no more specific zone inside
    // FieldPicker already claimed — see handlePageDrop's own WHY comment.
    <div onDragOver={handlePageDragOver} onDrop={handlePageDrop}>
      <AppShell>
        {metaError && (
        <p className="mb-6 rounded-xl border border-amber-300 bg-amber-50 px-4 py-3 text-amber-800">
          Couldn&apos;t load the field list from the API: {metaError}
        </p>
      )}

      {/* Two-pane layout matching Excel's own PivotTable UI: the grid is the main "worksheet"
          area (left, flexes to fill remaining width), the field picker is a docked pane (right,
          fixed width, sticky) — requested live, 2026-10-04 ("عايز اغير التصميم واخليه زي
          الإكسل"/"الصفحة كلها"), keeping this app's own gold/ink brand colors rather than
          switching to Excel's literal gray/white (confirmed same request). items-start: stops
          flexbox's default align-items:stretch from forcing both columns to the same height. */}
      <div className="flex items-start gap-6">
        {/* min-w-0: without it, a wide pivot grid can force this whole row (and the page) wider
            than the viewport — the classic flexbox min-width:auto trap. PivotGrid scrolls both axes
            internally (max-h-[80vh] overflow-auto on its own grid container, restored 2026-10-06
            after a page-scroll version of this briefly let a wide result bleed out under the sticky
            field-list sidebar — see PivotGrid.jsx's own WHY comment), so this just keeps the ROW
            from also trying to grow to fit it. */}
        <div className="min-w-0 flex-1">
          {/* No "Run query" button — matches Excel's own behavior (requested live, 2026-10-04,
              "فين زرار Run query في الإكسل؟"): the pivot re-queries itself automatically, see the
              debounced auto-run effect above. This is the only visible trace of that happening —
              Excel itself shows no persistent indicator either, just a brief status-bar flicker
              during a real recalculation. */}
          {loading && (
            <p className="mb-2 flex items-center gap-1.5 text-sm text-muted">
              <Loader2 className="h-3.5 w-3.5 animate-spin" />
              Updating…
            </p>
          )}

          {error && (
            <div className="mb-4 rounded-xl border border-red-300 bg-red-50 px-4 py-3 text-red-800">
              <p>{error}</p>
              <button type="button" onClick={() => runQuery()} disabled={loading} className="btn-secondary mt-2 disabled:cursor-not-allowed disabled:opacity-50">
                Retry
              </button>
            </div>
          )}

          {/* No top margin here (removed 2026-10-12, "الجزء دا عايز يطلع لفوق" — the Insert Slicer/
              Timeline/Calculated Field/Item ribbon) — with no loading/error message above it (the
              common case), this mt-6 was the only thing pushing it down from the column's own top,
              leaving it visibly lower than the Fields/zones cards next to it even though all three
              start at the same flex-row position. loading/error above already carry their own
              spacing (flex gap / mt-4) when they're actually shown. */}
          <div>
            <PivotToolbar
              request={request}
              onChange={setRequest}
              measureOptions={measures.map((m) => ({ value: m.field, label: m.displayName }))}
              dimensions={dimensions}
            />
            <PivotGrid
              result={result}
              rows={lastQuery?.rows}
              columns={lastQuery?.columns}
              rowFieldLabels={fieldLabels(lastQuery?.rows, dimensions)}
              columnFieldLabels={fieldLabels(hideMeasuresLabelIfSingleValue(lastQuery?.columns, lastQuery?.values?.length ?? 0), dimensions)}
              valueFields={displayValues}
              dimensions={dimensions}
              layout={reportLayout}
              repeatItemLabels={repeatItemLabels}
              fieldLayouts={fieldLayouts}
              // Real Excel's per-member OLAP drilldown (PLAN 8, EXPERIMENTAL) — keyed off `request`
              // (not lastQuery) like the Sort/Filter menus just below: clicking +/- edits the LIVE
              // expanded-member set and re-runs the query immediately (same auto-run effect as every
              // other field edit), it isn't a display-only toggle.
              rowsHierarchyDrilldown={request.rowsHierarchyDrilldown}
              drilldownSubtotals={request.subtotals}
              // Whether the member on this row is expanded — per outer item when the drilldown sits
              // after other Rows fields (Drilldown v3), else from the context-free list.
              isDrilldownMemberExpanded={(uniqueName, row) => {
                const current = request.rowsHierarchyDrilldown;
                if (!current) return false;
                const slot = row?.drill?.level ?? 0;
                if (row?.drill && slot > 0) {
                  const outer = outerKey(row.uniqueNames.slice(0, slot));
                  return (current.scopedExpansions ?? []).some((e) => outerKey(e.outerUniqueNames) === outer && e.memberUniqueName === uniqueName);
                }
                return current.expandedMemberUniqueNames.includes(uniqueName);
              }}
              onToggleDrilldownMember={(uniqueName, row) => {
                const current = request.rowsHierarchyDrilldown;
                if (!current) return;
                // Drilldown v3: combined with Rows fields BEFORE the slot, the expansion belongs to
                // the clicked row's outer item (real Excel expands per outer item) — stored in
                // scopedExpansions with the ancestry the backend reported in row.drill.
                const slot = row?.drill?.level ?? 0;
                if (row?.drill && slot > 0) {
                  const outerUniqueNames = row.uniqueNames.slice(0, slot);
                  const scoped = current.scopedExpansions ?? [];
                  const isThis = (e) => outerKey(e.outerUniqueNames) === outerKey(outerUniqueNames) && e.memberUniqueName === uniqueName;
                  const nextScoped = scoped.some(isThis)
                    ? scoped.filter((e) => !isThis(e))
                    : [...scoped, { outerUniqueNames, memberUniqueName: uniqueName, ancestors: row.drill.ancestorUniqueNames }];
                  setRequest({ ...request, rowsHierarchyDrilldown: { ...current, scopedExpansions: nextScoped } });
                  return;
                }
                // Like Excel, collapsing a member REMEMBERS the expansions under it (Drilldown v2,
                // 2026-10-14): only this member leaves the expanded list; its expanded descendants
                // stay, and the query just skips them while an ancestor is collapsed (see
                // visibleDrilldownExpansions). Expanding records the member's ancestry — the
                // backend's row.drill when combined with other fields, else the drilldown-alone
                // row's own uniqueNames (its full ancestry, see CubeQueryService.BuildDrilldownHeaders).
                const collapsing = current.expandedMemberUniqueNames.includes(uniqueName);
                const ancestors = { ...(current.expandedAncestors ?? {}) };
                if (!collapsing) {
                  if (row?.drill) {
                    ancestors[uniqueName] = row.drill.ancestorUniqueNames;
                  } else {
                    const found = (result?.rowHeaders ?? []).find((h) => (h.uniqueNames ?? []).includes(uniqueName));
                    const names = found?.uniqueNames ?? [];
                    ancestors[uniqueName] = names.slice(0, names.indexOf(uniqueName));
                  }
                }
                const expanded = collapsing
                  ? current.expandedMemberUniqueNames.filter((u) => u !== uniqueName)
                  : [...current.expandedMemberUniqueNames, uniqueName];
                setRequest({ ...request, rowsHierarchyDrilldown: { ...current, expandedMemberUniqueNames: expanded, expandedAncestors: ancestors } });
              }}
              // Columns drilldown (C4): its slot in the DISPLAYED columns (lastQuery), and +/- on the
              // header cells. Expansions are per outer item when real Columns fields sit before the
              // slot — their members come from the clicked header's own column (Σ Values skipped).
              columnsDrillSlot={lastQuery?.columnsHierarchyDrilldown ? lastQuery.columns.indexOf(lastQuery.columnsHierarchyDrilldown.topLevelField) : -1}
              isColumnDrilldownMemberExpanded={(uniqueName, column) => {
                const current = request.columnsHierarchyDrilldown;
                if (!current || !lastQuery) return false;
                const outer = columnsOuterIndexes(lastQuery.columns, current).map((i) => column?.uniqueNames?.[i]);
                return isDrilldownExpansionOn(current, uniqueName, outer);
              }}
              onToggleColumnDrilldownMember={(uniqueName, column, ancestors) => {
                const current = request.columnsHierarchyDrilldown;
                if (!current || !lastQuery) return;
                const outer = columnsOuterIndexes(lastQuery.columns, current).map((i) => column?.uniqueNames?.[i]);
                if (outer.some((u) => !u)) return;
                setRequest({ ...request, columnsHierarchyDrilldown: toggleDrilldownExpansion(current, uniqueName, outer, ancestors) });
              }}
              // request/onChange, not lastQuery: a header's Sort/Filter menu edits the LIVE field
              // config (same request/setRequest pair FieldPicker already uses) — rows/columns above
              // stay keyed off lastQuery so the field identifiers always match what's actually
              // displayed, but the edit itself targets the live request, same as every other
              // Sort/Filter control in FieldPicker (still needs Run query to take effect — this isn't
              // a special case, just a new entry point into the same existing flow).
              request={request}
              onChange={setRequest}
              onRefresh={runQuery}
              onHideFieldList={() => setShowFieldList(false)}
            />
          </div>

          {includeChart &&
            chartUnavailableReason(result, lastQuery) && (
              <p className="mt-4 rounded-xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm text-muted">{chartUnavailableReason(result, lastQuery)}</p>
            )}
          {includeChart && (
            <PivotChart result={result} valueFields={displayValues} chartType={chartType} rows={lastQuery?.rows} />
          )}

        </div>

        {/* sticky, no overflow/max-height: deliberately NOT giving this its own inner scroll —
            CSS2.1 §11.1.1 would force overflow-x to 'auto' too the moment overflow-y is set to
            anything but 'visible', silently clipping every Popover panel inside FieldPicker that
            extends past this pane's own width (several are min-w-[22rem], barely narrower than the
            pane itself) — Popover.jsx's own panel now portals to document.body so it wouldn't
            actually be clipped anymore, but adding a scrollbar here was never what was asked for
            regardless (tried 2026-10-12, reverted the SAME day: "انا ماقولتلكش زود scroll bar!!") —
            the actual ask is for the content to genuinely FIT without needing to scroll anywhere,
            not to relocate the scrollbar from the page onto this pane instead. Plain sticky avoids
            the CSS2.1 clipping rule entirely since position:sticky sets no overflow property.
            top-8 matches AppShell's own p-8 page padding. */}
        {/* Narrowed back down (640 → 600 → 460): real Excel's own Field List is a fixed, fairly
            narrow pane regardless of monitor size — the WORKSHEET (here, PivotGrid) is what's
            supposed to dominate the page, not the field picker (flagged live, 2026-10-04: "المفروض
            ان المساحة الكبرى تبقى للداتا"). The two-column field-list/zones layout inside
            FieldPicker now just runs tighter in less width instead of getting more room. */}
        {showFieldList ? (
          <aside className="sticky top-8 w-[460px] shrink-0">
            <FieldPicker
              value={request}
              onChange={setRequest}
              dimensions={dimensions}
              measures={measures}
              tables={tables}
              tablesLoading={tablesLoading}
              enabledTableNames={enabledTableNames}
              onEnabledTableNamesChange={updateEnabledTables}
              fieldsLoading={fieldsLoading}
              onReset={handleReset}
              fieldLayouts={fieldLayouts}
              reportLayout={reportLayout}
              repeatItemLabels={repeatItemLabels}
              onFieldLayoutsChange={setFieldLayouts}
            />
            {/* Moved here from the main column (2026-10-12, "تحت الفلاتر و الحته بتاعت اختيار ال
                dimensions") — "الفلاتر" meant FieldPicker's OWN Filters zone (one of its Rows/
                Columns/Values/Filters drop zones above), not the Report Filter pills PivotGrid
                renders over the table; "below that AND the dimension-selection part" together just
                means below this whole FieldPicker sidebar, not somewhere in the main content column
                at all (an earlier placement right above the grid was corrected after guessing wrong
                once already — see PivotBuilder's own git history for that one-day detour). Stacked
                vertically (flex-col, w-full inputs) instead of the old wide-row flex-wrap layout,
                which only made sense stretched across the main column — unmodified in this narrow
                460px sidebar, that row would just wrap into a messy multi-line block. */}
            {result && (
              // p-3 (not .card's own p-6) and gap-1.5 (not gap-3), plus py-2 on both text inputs
              // below (not .input-field's own py-3) — shrunk 2026-10-12 ("صغر ارتفاع دي شوية", then
              // cut further the same day chasing a whole-page scrollbar — "انت مش راضي تحل
              // المشكلة"): this card sits right under the zones column in a tight sidebar, so the
              // more generous default padding/spacing other `.card`/`.input-field` surfaces use
              // elsewhere in the app was more height than this particular spot could spare.
              <div className="card mt-2 flex flex-col gap-1.5 p-3">
                <input
                  type="text"
                  list="snapshot-name-list"
                  className="input-field w-full py-2"
                  placeholder="Snapshot name…"
                  value={snapshotName}
                  onChange={(e) => setSnapshotName(e.target.value)}
                  title="Pick an existing snapshot to save a new version under it, or type a new name"
                />
                {/* The native combo behavior itself (suggestions while still accepting free text) —
                    no custom dropdown component needed for "pick existing or type new." */}
                <datalist id="snapshot-name-list">
                  {existingSnapshotNames.map((name) => (
                    <option key={name} value={name} />
                  ))}
                </datalist>
                {/* Only offered once a Snapshot name exists — nothing to attach a version label TO
                    before that — and only shown as a text field once the user actually asks for one,
                    not stacked underneath by default (see showVersionName's own WHY comment above). */}
                {snapshotName.trim() && !showVersionName && (
                  <button
                    type="button"
                    onClick={() => setShowVersionName(true)}
                    className="inline-flex items-center gap-1 self-start text-xs font-medium text-gold-strong hover:underline"
                  >
                    <Plus className="h-3 w-3" />
                    Add a version name
                  </button>
                )}
                {showVersionName && (
                  <input
                    type="text"
                    autoFocus
                    className="input-field w-full py-2"
                    placeholder="Version name (optional)…"
                    value={versionName}
                    onChange={(e) => setVersionName(e.target.value)}
                    title="An optional label for just this version, e.g. &quot;Q3 Final&quot;"
                  />
                )}
                <div className="flex gap-2">
                  <select
                    className="w-1/2 rounded-xl border border-slate-200 bg-white/90 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent"
                    value={tableStyle}
                    onChange={(e) => setTableStyle(e.target.value)}
                    title="Table style for the saved snapshot"
                  >
                    <option value="Default">Style: Default</option>
                    <option value="Banded">Style: Banded rows</option>
                    <option value="Dark">Style: Dark</option>
                    <option value="Minimal">Style: Minimal</option>
                  </select>
                  <select
                    className="w-1/2 rounded-xl border border-slate-200 bg-white/90 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent"
                    value={reportLayout}
                    onChange={(e) => setReportLayout(e.target.value)}
                    title="Report layout for the Rows area"
                  >
                    <option value="Tabular">Layout: Tabular</option>
                    <option value="Compact">Layout: Compact</option>
                    <option value="Outline">Layout: Outline</option>
                  </select>
                </div>
                {/* Real Excel grays this option out entirely for Compact Form (no separate per-level
                    column to repeat a label into) — disabled here the same way, rather than letting it
                    silently have no effect (see PivotGrid.jsx's own WHY comment). */}
                <label
                  className={`flex items-center gap-2 text-sm ${reportLayout === "Compact" ? "text-muted" : "text-ink"}`}
                  title="Repeat every row field's label on every row, instead of showing it once per group"
                >
                  <input
                    type="checkbox"
                    checked={repeatItemLabels}
                    disabled={reportLayout === "Compact"}
                    onChange={(e) => setRepeatItemLabels(e.target.checked)}
                  />
                  Repeat All Item Labels
                </label>
                <label className="flex items-center gap-2 text-sm text-ink">
                  <input type="checkbox" checked={includeChart} onChange={(e) => setIncludeChart(e.target.checked)} />
                  Include chart
                </label>
                {includeChart && (
                  <select
                    className="w-full rounded-xl border border-slate-200 bg-white/90 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent"
                    value={chartType}
                    onChange={(e) => setChartType(e.target.value)}
                    title="Chart type"
                  >
                    <option value="Bar">Chart: Bar</option>
                    <option value="Line">Chart: Line</option>
                    <option value="Pie">Chart: Pie</option>
                  </select>
                )}
                <div className="flex gap-2">
                  <button
                    type="button"
                    onClick={() => handleSave()}
                    disabled={saving || !snapshotName.trim()}
                    className="btn-primary flex-1 disabled:cursor-not-allowed disabled:opacity-50"
                  >
                    {saving ? "Saving…" : "Save Snapshot"}
                  </button>
                  <button
                    type="button"
                    onClick={handleExport}
                    disabled={exporting}
                    className="btn-secondary flex-1 inline-flex items-center justify-center gap-1.5 disabled:cursor-not-allowed disabled:opacity-50"
                  >
                    <FileSpreadsheet className="h-4 w-4" style={{ color: "#217346" }} />
                    {exporting ? "Exporting…" : "Export to Excel"}
                  </button>
                </div>
                {saveMessage && <span className="text-sm font-medium text-green-700">{saveMessage}</span>}
                {saveError && <span className="text-sm font-medium text-red-700">{saveError}</span>}
                {exportError && <span className="text-sm font-medium text-red-700">{exportError}</span>}
              </div>
            )}
          </aside>
        ) : (
          // Excel's own "Hide Field List" (ribbon toggle + value-cell context menu item) hides
          // this whole pane to give the worksheet more room — this button is the only way back,
          // same idea as Excel's own re-toggle, just as a visible button instead of a ribbon state.
          <button
            type="button"
            onClick={() => setShowFieldList(true)}
            className="btn-secondary sticky top-8 h-fit shrink-0 px-3 py-2 text-xs"
          >
            Show Field List
          </button>
        )}
      </div>
      </AppShell>
    </div>
  );
}
