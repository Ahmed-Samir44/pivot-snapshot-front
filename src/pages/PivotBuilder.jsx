import { useEffect, useRef, useState } from "react";
import { FileSpreadsheet, PanelRightOpen, Plus, Save, X } from "lucide-react";
import Popover from "../components/forms/Popover";
import MessageBar from "../components/layout/MessageBar";
import Toasts from "../components/layout/Toasts";
import { isCubeTemporarilyDown } from "../services/apiError";
import { useRetryCountdown } from "../utils/useRetryCountdown";
import { useUndoHistory } from "../utils/useUndoHistory";
import AppShell from "../components/layout/AppShell";
import FieldPicker from "../components/pivot/FieldPicker";
import PivotToolbar from "../components/pivot/PivotToolbar";
import PivotGrid from "../components/pivot/PivotGrid";
import PivotTableOptionsDialog from "../components/pivot/PivotTableOptionsDialog";
import StatusBar from "../components/pivot/StatusBar";
import { ZOOM_LEVELS } from "../utils/zoom";
import PaneSplitter from "../components/layout/PaneSplitter";
import { DEFAULT_PIVOT_OPTIONS } from "../utils/pivotOptions";
import { onePerField } from "../utils/filterKinds";
import { isFilterComplete } from "../utils/filterModes";
import { formatDimensionName } from "../utils/cubeMeta";
import PivotChart from "../components/pivot/PivotChart";
import { chartUnavailableReason } from "../utils/chartLayout";
import { queryPivot } from "../services/pivotApi";
import { getTables, getDimensions, getDimensionsForTable, getMeasures, getMembers } from "../services/cubeMetaApi";
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
  // Defaults to Top: Excel shows subtotals at the top of each group as soon as a second Rows field
  // is added (confirmed by the user against Excel, 2026-10-18 — it used to default to Off). With
  // one field there are no groups, so nothing shows (the backend doesn't block Running Total/Rank
  // then either). A pivot saved with its own choice keeps it.
  subtotals: "Top",
  // PivotTable Options > Totals & Filters > "Include filtered items in totals" (2026-10-19) — off is
  // Excel's default. Part of `request` (unlike pivotOptions) since it changes the query's totals.
  includeFilteredItemsInTotals: false,
  // "Allow multiple filters per field" (same tab) — off is Excel's default: one filter per field.
  allowMultipleFiltersPerField: false,
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
  // Excel's Group Selection on text items: { field, name, members } (see GroupItemsPanel).
  itemGroups: [],
  showItemsWithNoData: [],
  // Real Excel's per-field "Show items with no data" (PLAN 7, 2026-10-13, EXPERIMENTAL — see the
  // plan file's own WHY comment on why this needs live-cube verification). Unlike showItemsWithNoData
  // above (per-AXIS), this is a plain list of field names — a field in this list keeps its own
  // zero-data members even when the rest of its axis still filters normally. Query-affecting (unlike
  // fieldLayouts' Compact/RepeatItemLabels/CustomName/InsertBlankLine), so it stays inside `request`
  // and IS sent to queryPivot, not stripped out. Rows-only; any number of fields (2026-10-20 — each
  // gets its own orphan branch, see MdxPivotQueryBuilder).
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
// Name") for the grid's header row, instead of leaving it blank. A field whose table isn't loaded in
// the field list (collapsed, or restored from a saved pivot) is formatted from its own identifier
// instead of shown raw ("[Dim X].[Y].[Y]" — seen 2026-10-19), the same way Snapshot History labels it.
// fieldLayouts: a field's Field Settings › Custom Name wins, as in Excel (2026-10-21, for Columns).
function fieldLabels(fieldIds, dimensions, fieldLayouts = []) {
  if (!fieldIds) return [];
  return fieldIds.map((id) =>
    id === MEASURES_PLACEHOLDER
      ? "Σ Values"
      : (fieldLayouts.find((f) => f.field === id)?.customName || (dimensions.find((d) => d.field === id)?.displayName ?? formatDimensionName(id))),
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

// Workspace view settings, per browser (2026-10-20 UI refresh) — conveniences only: unreadable
// storage just means the defaults.
const FIELD_PANE_WIDTH_KEY = "pivotSnapshot.fieldPaneWidth";
const ZOOM_KEY = "pivotSnapshot.gridZoom";
const FIELD_PANE_DEFAULT_WIDTH = 460;
const FIELD_PANE_MIN_WIDTH = 380;
const FIELD_PANE_MAX_WIDTH = 760;

function readStoredNumber(key, fallback, min, max) {
  try {
    const value = Number(localStorage.getItem(key));
    return localStorage.getItem(key) !== null && Number.isFinite(value) ? Math.min(max, Math.max(min, value)) : fallback;
  } catch {
    return fallback;
  }
}

function writeStored(key, value) {
  try {
    localStorage.setItem(key, String(value));
  } catch {
    // Not persisted — the setting still applies for this page view.
  }
}

// A filter is only ready to send once whichever fields ITS mode actually needs are filled in —
// each mode uses a different subset of PivotFilter's optional fields (see the backend's
// MdxPivotQueryBuilder.ValidateFilters for the authoritative per-mode rules this mirrors).
// (2026-10-20: GreaterThan/LessThan/Between used to fall through to the member-list check and were
// silently DROPPED from every query — a value filter has no members. isRuleFilterComplete knows every mode.)
// Excel only shows Σ Values in an area when there are 2+ value fields; with one it isn't there, so a
// field dropped next to it effectively comes first. Keeping Σ Values LAST on its axis in that case
// gives Excel's layout (the value caption under the column items, not a band above them) and keeps
// a new field from landing behind it (2026-10-21).
function measuresLastWhenSingleValue(request) {
  if (request.values.length > 1) return request;
  // Moving Σ Values shifts the levels after it down by one — per-level sorts follow their field.
  const move = (fields, levelSorts) => {
    const at = fields.indexOf(MEASURES_PLACEHOLDER);
    if (at < 0 || at === fields.length - 1) return null;
    return {
      fields: [...fields.filter((f) => f !== MEASURES_PLACEHOLDER), MEASURES_PLACEHOLDER],
      levelSorts: levelSorts.map((ls) => (ls.level > at ? { ...ls, level: ls.level - 1 } : ls)),
    };
  };
  const rows = move(request.rows, request.rowLevelSorts);
  const columns = move(request.columns, request.columnLevelSorts);
  if (!rows && !columns) return request;
  return {
    ...request,
    ...(rows ? { rows: rows.fields, rowLevelSorts: rows.levelSorts } : {}),
    ...(columns ? { columns: columns.fields, columnLevelSorts: columns.levelSorts } : {}),
  };
}

// (isFilterComplete itself lives in utils/filterModes.js — Show Details uses it too.)

// Everything in `request` EXCEPT each value's `format`/`conditionalFormat` — those two are pure
// display settings (see displayValues' own WHY comment below) that must NOT mark the on-screen
// result as stale, unlike every other field here (rows, columns, filters, sort, groupings,
// totals, showItemsWithNoData, calculatedFields/Items, and even a value's
// showValuesAs — that one DOES need a fresh query since it's computed server-side). Used to detect
// "the user changed something that actually invalidates the table currently on screen."
// slicerFields/timelineFields are left out too (2026-10-25): they only say which widgets are on the
// sheet (the backend never reads them) — inserting a slicer re-ran the whole query, which Excel
// doesn't do; a slicer's actual selection lives in `filters` and still refreshes the table.
function structuralSignature(req) {
  const { slicerFields: _slicers, timelineFields: _timelines, ...rest } = req;
  return JSON.stringify({
    ...rest,
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
      const next = measuresLastWhenSingleValue(normalizeMeasuresPlacement(typeof updater === "function" ? updater(prev) : updater));
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
        // The sort goes on the axis's first REAL field — level 0 can be Σ Values, which can't be sorted
        // (2026-10-21, found by the feature test: dragging a field into Columns after Σ Values failed
        // with "ColumnLevelSorts references level 0, which is where [Measures] sits").
        const firstRealLevel = (fields) => fields.findIndex((f) => f !== MEASURES_PLACEHOLDER);
        const rowsJustGotFirstField = prev.rows.filter((f) => f !== MEASURES_PLACEHOLDER).length === 0 && firstRealLevel(withDefaultSort.rows) >= 0;
        // Not for a drilldown slot (Drilldown v3: starting a drilldown on empty Rows makes it the
        // first Rows field) — a per-level sort can't apply to a drilldown axis (the backend rejects
        // it); the hierarchy keeps its own natural order unless sorted from its header.
        if (rowsJustGotFirstField && withDefaultSort.sort?.axis !== "Rows" && withDefaultSort.rowLevelSorts.length === 0 && !withDefaultSort.rowsHierarchyDrilldown) {
          withDefaultSort = { ...withDefaultSort, rowLevelSorts: [{ level: firstRealLevel(withDefaultSort.rows), direction: "Ascending", byMeasureField: null }] };
        }
        const prevColumnFieldCount = prev.columns.filter((f) => f !== MEASURES_PLACEHOLDER).length;
        const nextColumnFieldCount = withDefaultSort.columns.filter((f) => f !== MEASURES_PLACEHOLDER).length;
        const columnsJustGotFirstField = prevColumnFieldCount === 0 && nextColumnFieldCount > 0;
        // Same exception as Rows for a Columns drilldown slot (C4).
        if (columnsJustGotFirstField && withDefaultSort.sort?.axis !== "Columns" && withDefaultSort.columnLevelSorts.length === 0 && !withDefaultSort.columnsHierarchyDrilldown) {
          withDefaultSort = { ...withDefaultSort, columnLevelSorts: [{ level: firstRealLevel(withDefaultSort.columns), direction: "Ascending", byMeasureField: null }] };
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
  // The last query failed only because the data server is busy/unreachable (ApiError code) — shown as
  // a calm notice that retries by itself (2026-10-20) instead of a red error. queryFailures restarts
  // the countdown after each failed attempt.
  const [errorTemporary, setErrorTemporary] = useState(false);
  const [queryFailures, setQueryFailures] = useState(0);
  const [loading, setLoading] = useState(false);
  const [measures, setMeasures] = useState([]);
  const [metaError, setMetaError] = useState(null);
  // Same for the field list: a busy/unreachable server reloads it on its own (metaAttempt re-runs
  // the loading effects below).
  const [metaTemporary, setMetaTemporary] = useState(false);
  const [metaAttempt, setMetaAttempt] = useState(0);
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
  // The cube's full field list (loaded in the background, see its effect below) — declared here because
  // the placed-fields lookup just below uses it.
  const [allDimensions, setAllDimensions] = useState(null);
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
    ...request.slicerFields,
    ...request.timelineFields,
  ]);
  // From every table fetched so far AND the full field list once it has loaded (2026-10-20, found by
  // the feature test): a field placed from a table never expanded this session — a saved pivot's
  // filter, a slicer or timeline added from the whole cube — keeps its name and date/number kind.
  const knownFields = new Map([...Object.values(fieldsByTable).flat(), ...(allDimensions ?? [])].map((d) => [d.field, d]));
  const stillPlacedButCollapsed = [...knownFields.values()].filter((d) => placedFieldIds.has(d.field) && !enabledFieldIds.has(d.field));
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
  // Excel's PivotTable Options > Layout & Format (2026-10-18): name, merged labels, empty/error cell
  // text, report filter area layout. Display-only, so — like fieldLayouts — a sibling of `request`,
  // never part of structuralSignature (a change never re-queries the cube). Merged with the
  // defaults so a state saved before a newer option existed still gets it.
  const [pivotOptions, setPivotOptions] = useState(() => ({ ...DEFAULT_PIVOT_OPTIONS, ...(readPersistedBuilderState()?.pivotOptions ?? {}) }));
  // Formats set on individual cells with Home › Number (utils/cellFormats.js) — display-only, kept
  // with the builder state and every saved version; the grid reports the selection and applies them
  // (gridCommandRef).
  const [cellFormats, setCellFormats] = useState(() => readPersistedBuilderState()?.cellFormats ?? {});
  const gridCommandRef = useRef(null);

  const [showPivotOptions, setShowPivotOptions] = useState(false);
  // Excel's value-cell context menu has "Hide Field List" — toggles the FIELDS sidebar's own
  // visibility, with a small floating button to bring it back (Excel keeps the same toggle on the
  // PivotTable Analyze ribbon too, not just the context menu, but there's no ribbon button for it
  // here yet — this is the only entry point so far).
  const [showFieldList, setShowFieldList] = useState(true);
  // Workspace view state (2026-10-20 UI refresh), per browser: the Fields pane's dragged width and
  // the grid's zoom. The selected cells' Sum/Average/Count come up from PivotGrid for the status bar.
  const [fieldPaneWidth, setFieldPaneWidth] = useState(() => readStoredNumber(FIELD_PANE_WIDTH_KEY, FIELD_PANE_DEFAULT_WIDTH, FIELD_PANE_MIN_WIDTH, FIELD_PANE_MAX_WIDTH));
  const [zoom, setZoom] = useState(() => {
    const stored = readStoredNumber(ZOOM_KEY, 1, ZOOM_LEVELS[0], ZOOM_LEVELS[ZOOM_LEVELS.length - 1]);
    return ZOOM_LEVELS.includes(stored) ? stored : 1;
  });
  const [selectionSummary, setSelectionSummary] = useState(null);
  // The Fields pane header element, where FieldPicker puts its gear (like Excel’s task pane).
  const [paneToolsEl, setPaneToolsEl] = useState(null);
  useEffect(() => writeStored(FIELD_PANE_WIDTH_KEY, fieldPaneWidth), [fieldPaneWidth]);
  useEffect(() => writeStored(ZOOM_KEY, zoom), [zoom]);
  // Excel's "Defer Layout Update" (2026-10-20, bottom of the Field List): while checked, field list
  // edits collect in pendingRequest (what the field list shows) and nothing re-queries; "Update"
  // applies them in one go. Unchecking applies whatever is pending. Not persisted, like Excel's.
  const [deferLayout, setDeferLayout] = useState(false);
  const [pendingRequest, setPendingRequest] = useState(null);

  // Excel's Undo / Redo (Ctrl+Z / Ctrl+Y, and the Home tab's buttons) over everything that describes
  // the PivotTable (2026-10-21). Restored with setRequestRaw so the state comes back exactly, not
  // re-normalized (setRequest adds default sorts on an axis's first field).
  const history = useUndoHistory(
    { request, reportLayout, repeatItemLabels, fieldLayouts, pivotOptions, cellFormats },
    (snapshot) => {
      setPendingRequest(null);
      setRequestRaw(snapshot.request);
      setReportLayout(snapshot.reportLayout);
      setRepeatItemLabels(snapshot.repeatItemLabels);
      setFieldLayouts(snapshot.fieldLayouts);
      setPivotOptions(snapshot.pivotOptions);
      setCellFormats(snapshot.cellFormats);
    },
  );
  useEffect(() => {
    const onKey = (e) => {
      if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
      const k = e.key.toLowerCase();
      if (k !== "z" && k !== "y") return;
      // Typing in a box keeps the browser's own text undo.
      if (e.target.closest?.("input, textarea, select, [contenteditable='true']")) return;
      e.preventDefault();
      if (k === "y" || e.shiftKey) history.redo();
      else history.undo();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  });
  const applyPendingLayout = () => {
    if (pendingRequest) setRequest(pendingRequest);
    setPendingRequest(null);
  };
  // While deferred changes are pending, EVERY editor (ribbon, grid menus, options dialog, page drop)
  // edits that pending copy too — editing `request` instead would be overwritten by Update, or
  // undone, since the pending copy was taken before the edit (found in review 2026-10-20).
  const editingRequest = pendingRequest ?? request;
  const editRequest = (updater) => {
    if (!pendingRequest) {
      setRequest(updater);
      return;
    }
    setPendingRequest((prev) => normalizeMeasuresPlacement(typeof updater === "function" ? updater(prev) : updater));
  };
  const [saving, setSaving] = useState(false);
  const [saveMessage, setSaveMessage] = useState(null);
  const [saveError, setSaveError] = useState(null);
  const [exporting, setExporting] = useState(false);
  const [exportError, setExportError] = useState(null);
  // Why the last export is plain values instead of an Excel PivotTable (see excelPivot.js), or null.
  const [exportNotice, setExportNotice] = useState(null);

  // Every field of every table, for the Field List's search (Excel's search matches FIELD names,
  // not just table names — found by the full-app test 2026-10-18). Fetched once in the background:
  // the first read takes ~40 s on this cube, after which the backend serves it from its cache. Until
  // it arrives (or if it fails), search still matches table names and already-loaded fields.
  useEffect(() => {
    getDimensions()
      .then(setAllDimensions)
      .catch(() => {});
  }, []);

  useEffect(() => {
    Promise.all([getTables(), getMeasures()])
      .then(([tbls, meas]) => {
        setTables(tbls);
        setMeasures(meas);
        setMetaError(null);
        setMetaTemporary(false);
      })
      .catch((err) => {
        setMetaTemporary(isCubeTemporarilyDown(err));
        // A rejection here isn't always a plain Error with a real .message (caught live,
        // 2026-09-27: a concurrent-token-refresh race — see dataverseAuth.js's WHY comment —
        // produced one whose falsy .message made {metaError && (...)} render nothing at all,
        // a completely silent failure with no visible error anywhere). Falling back through
        // toString()/JSON.stringify guarantees something is always shown here.
        setMetaError(String(err?.message || err?.toString?.() || JSON.stringify(err) || "Unknown error."));
      })
      .finally(() => setTablesLoading(false));
  }, [metaAttempt]);

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
        setMetaTemporary(isCubeTemporarilyDown(err));
        setMetaError(String(err?.message || err?.toString?.() || JSON.stringify(err) || "Unknown error."));
      })
      .finally(() => setFieldsLoading(false));
  }, [enabledTableNames, fieldsByTable, tables, metaAttempt]);

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
        includeFilteredItemsInTotals: Boolean(request.includeFilteredItemsInTotals),
        allowMultipleFiltersPerField: Boolean(request.allowMultipleFiltersPerField),
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
        // Group Selection: only for a field still on an axis, not date/number-grouped or drilled
        // down (MdxPivotQueryBuilder.ValidateItemGroups) — a stale group is dropped, not an error.
        itemGroups: (request.itemGroups ?? []).filter(
          (g) =>
            (request.rows.includes(g.field) || request.columns.includes(g.field)) &&
            g.members.length > 0 &&
            !request.dateGroupings.some((d) => d.field === g.field) &&
            !request.numericGroupings.some((n) => n.field === g.field) &&
            request.rowsHierarchyDrilldown?.topLevelField !== g.field &&
            request.columnsHierarchyDrilldown?.topLevelField !== g.field,
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
      setErrorTemporary(isCubeTemporarilyDown(err));
      if (isCubeTemporarilyDown(err)) setQueryFailures((n) => n + 1);
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
        JSON.stringify({ request, tableStyle, includeChart, chartType, reportLayout, repeatItemLabels, fieldLayouts, pivotOptions, cellFormats }),
      );
    } catch {
      // Per-browser convenience only — same no-op-on-failure as every other localStorage write in
      // this file; worst case, the next visit just starts blank again.
    }
  }, [request, tableStyle, includeChart, chartType, reportLayout, repeatItemLabels, fieldLayouts, pivotOptions, cellFormats]);

  // Orphaned-entry cleanup for fieldLayouts, same reasoning as request.fieldSubtotals' own cleanup
  // in FieldPicker.jsx (moveField/addHierarchyToAxis) — a field that leaves Rows should take its
  // Layout override with it. Done here instead of inside FieldPicker's onChange handlers because
  // fieldLayouts lives OUTSIDE `request` (see its own WHY comment above), so FieldPicker's
  // request-only onChange has no way to touch it directly.
  // Columns fields count too since their Field Settings gained Custom Name (2026-10-25) — Rows-only,
  // this wiped a column field's custom name on every reload.
  useEffect(() => {
    setFieldLayouts((prev) => prev.filter((f) => request.rows.includes(f.field) || request.columns.includes(f.field)));
  }, [request.rows, request.columns]);

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
  // Excel's PivotTable Analyze › Clear › Clear All (2026-10-20, "i need the system to match excel" —
  // replaces the field pane's "Reset", which Excel doesn't have). Like Excel's, it empties the
  // PivotTable — every field off every area, every filter, sort, slicer/timeline, grouping and
  // per-field setting — but keeps the PivotTable itself: its name and PivotTable Options, the report
  // layout, the Design tab's Grand Totals/Subtotals, and the calculated fields defined for it (they
  // stay in the field list, as in Excel). It applies at once even under Defer Layout Update, like
  // Excel's ribbon command. The snapshot's own name/style/chart settings aren't part of the
  // PivotTable, so they're left alone too.
  // Excel's "click here to cancel" on a long refresh (2026-10-21): stops waiting for the running query
  // (the server cancels it on the cube too) and offers Retry.
  const handleCancelQuery = () => {
    queryTokenRef.current += 1;
    queryAbortControllerRef.current?.abort();
    setLoading(false);
    setErrorTemporary(false);
    setError("The query was cancelled.");
  };

  const handleClearAll = () => {
    queryTokenRef.current += 1;
    queryAbortControllerRef.current?.abort();
    setLoading(false);
    setPendingRequest(null);
    setRequest((prev) => ({
      ...EMPTY_REQUEST,
      grandTotals: prev.grandTotals,
      subtotals: prev.subtotals,
      includeFilteredItemsInTotals: prev.includeFilteredItemsInTotals,
      allowMultipleFiltersPerField: prev.allowMultipleFiltersPerField,
      calculatedFields: prev.calculatedFields,
    }));
    setFieldLayouts([]);
    setCellFormats({});
    setLastQuery(null);
    setResult(null);
    setError(null);
    setExportError(null);
    setExportNotice(null);
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
    editRequest((prev) => removeFieldEverywhere(prev, payload.field));
  };

  // Returns true when the version was saved (the Save popover closes on it).
  const handleSave = async () => {
    if (!snapshotName.trim() || !result || !lastQuery) {
      return false;
    }
    setSaving(true);
    setSaveError(null);
    setSaveMessage(null);
    try {
      // The SAME bundle shape (and the SAME live `request`, not the cleaned/filtered one Run Query
      // sends) the persistence effect above already writes to localStorage — reused as-is here so
      // Edit later restores this exactly as it looked while being edited, not just a filtered/
      // query-ready subset of it. See BUILDER_STATE_STORAGE_KEY's own WHY comment.
      const builderStateJson = JSON.stringify({ request, tableStyle, includeChart, chartType, reportLayout, repeatItemLabels, fieldLayouts, pivotOptions, cellFormats });
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
        pivotOptions,
      );
      setSaveMessage(`Saved as version ${saved.versionNumber}.`);
      // A successful save means this snapshot name now definitely exists — refreshed so the NEXT
      // save's datalist offers it too, without needing a page reload (e.g. saving a brand-new name
      // for the first time).
      if (!existingSnapshotNames.includes(snapshotName.trim())) {
        setExistingSnapshotNames((prev) => [...prev, snapshotName.trim()]);
      }
      return true;
    } catch (err) {
      // SnapshotTooLarge included — a flat, non-overridable ceiling now (removed 2026-09-27 on
      // request), so this just shows like any other save error rather than a confirm-and-retry.
      setSaveError(err.message);
      return false;
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
      // File: the snapshot name when one's typed, else the PivotTable Name (2026-10-20, was "Pivot").
      const exported = await exportPivotToExcel({ ...exportOptionsFor(lastQuery, result), filename: snapshotName.trim() || pivotOptions.name?.trim() || "Pivot" });
      setExportNotice(exported.pivotTable ? null : `Exported as values, not a PivotTable: ${exported.reason}.`);
    } catch (err) {
      setExportError(err.message);
    } finally {
      setExporting(false);
    }
  };

  // What exportPivotToExcel needs to draw one query's result like the live grid — shared by Export
  // and Show Report Filter Pages.
  const exportOptionsFor = (query, queryResult) => ({
    result: queryResult,
    rows: query.rows,
    columns: query.columns,
    rowFieldLabels: fieldLabels(query.rows, dimensions),
    columnFieldLabels: fieldLabels(hideMeasuresLabelIfSingleValue(query.columns, query.values.length), dimensions, fieldLayouts),
    valueFields: displayValues,
    // Same layout as the live grid, so the export draws the same rows.
    layout: reportLayout,
    repeatItemLabels,
    fieldLayouts,
    // PivotTable Options (empty/error text, merged labels) and the Report Filter area above the
    // table. Every applied filter is listed, slicer/timeline ones too — the file has no slicer
    // widget to show them (same as the saved snapshot).
    pivotOptions,
    cellFormats,
    filters: query.filters,
    filterFieldLabels: Object.fromEntries(query.filters.map((f) => [f.field, fieldLabels([f.field], dimensions)[0]])),
    // A drilldown alone keeps its single ragged column; combined with other Rows fields
    // (Drilldown v3) it's a normal field slot, indented by member depth.
    ...(() => {
      const d = query.rowsHierarchyDrilldown;
      if (!d) return {};
      const others = query.rows.filter((f) => f !== MEASURES_PLACEHOLDER && f !== d.topLevelField);
      return others.length === 0 ? { drilldown: true } : { drillSlot: query.rows.indexOf(d.topLevelField) };
    })(),
    drilldownSubtotals: query.subtotals,
    // Columns drilldown (C4): header rows from buildColumnHeaderRows, same as the grid.
    columnsDrillSlot: query.columnsHierarchyDrilldown ? query.columns.indexOf(query.columnsHierarchyDrilldown.topLevelField) : -1,
  });

  // Excel's "Show Report Filter Pages" (2026-10-20): one copy of the pivot per item of a Report Filter
  // field, each filtered to that item. Excel makes worksheets — here, sheets of one downloaded
  // workbook. One query per item, one at a time (the shared cube), capped at MAX_FILTER_PAGES.
  // Raised 50 → 100 (2026-10-20); each page is one query on the shared cube, run one at a time.
  const MAX_FILTER_PAGES = 100;
  // The Report Filter area's fields: filters on fields not on Rows/Columns, minus slicer/timeline
  // fields (the grid draws those as widgets, not in the Report Filter area).
  const reportFilterFieldsOf = (req) =>
    [...new Set(req.filters.map((f) => f.field).filter(Boolean))].filter(
      (f) => !req.rows.includes(f) && !req.columns.includes(f) && !req.slicerFields.includes(f) && !req.timelineFields.includes(f),
    );
  const [filterPagesRunning, setFilterPagesRunning] = useState(false);
  const [filterPagesStatus, setFilterPagesStatus] = useState(null);
  const handleShowReportFilterPages = async (field) => {
    if (!lastQuery || filterPagesRunning) return;
    setExportError(null);
    setFilterPagesRunning(true);
    try {
      const existing = lastQuery.filters.find((f) => f.field === field);
      let members = (await getMembers(field)).filter((m) => m !== "");
      // A field already narrowed to some items: one page per CHOSEN item, like Excel (or every item
      // but the excluded ones).
      if (existing?.mode === "Members" && existing.includedMembers.length > 0) members = members.filter((m) => existing.includedMembers.includes(m));
      if (existing?.mode === "Exclude") members = members.filter((m) => !existing.includedMembers.includes(m));
      const truncated = members.length > MAX_FILTER_PAGES;
      members = members.slice(0, MAX_FILTER_PAGES);
      const { exportPivotToExcel, createWorkbook, downloadWorkbook } = await import("../utils/exportExcel");
      const workbook = createWorkbook();
      let pages = 0;
      for (const [i, member] of members.entries()) {
        setFilterPagesStatus(`Building page ${i + 1} of ${members.length}…`);
        const pageQuery = {
          ...lastQuery,
          filters: [...lastQuery.filters.filter((f) => f.field !== field), { field, mode: "Members", includedMembers: [member] }],
        };
        const pageResult = await queryPivot(pageQuery);
        // No data for this item — no page, like Excel. (A Values-only pivot always has its one row,
        // so emptiness is judged by the values.)
        if (pageResult.rowHeaders.length === 0 || pageResult.cells.every((row) => row.every((v) => v === null || v === undefined))) continue;
        await exportPivotToExcel({ ...exportOptionsFor(pageQuery, pageResult), workbook, sheetName: member, asPivotTable: false });
        pages++;
      }
      if (pages === 0) throw new Error("None of the items has data.");
      await downloadWorkbook(workbook, `${snapshotName.trim() || pivotOptions.name?.trim() || "Pivot"} - ${fieldLabels([field], dimensions)[0]}`);
      setFilterPagesStatus(`${pages} page${pages === 1 ? "" : "s"} exported${truncated ? ` (first ${MAX_FILTER_PAGES} items only)` : ""}.`);
    } catch (err) {
      setFilterPagesStatus(null);
      setExportError(err.message);
    } finally {
      setFilterPagesRunning(false);
    }
  };

  // Excel's Analyze › Clear: Clear All empties the pivot (the same as Reset); Clear Filters drops
  // every filter, leaving a slicer/timeline field with an "(All)" entry so the widget stays.
  // A Report Filter field (and a slicer's/timeline's) goes back to "(All)" rather than leaving — like
  // Excel, Clear Filters keeps the fields where they are (it used to drop the Filters-area fields).
  const handleClearFilters = () =>
    editRequest((prev) => ({
      ...prev,
      filters: prev.filters
        .filter((f) => !prev.rows.includes(f.field) && !prev.columns.includes(f.field))
        .filter((f, i, all) => all.findIndex((x) => x.field === f.field) === i)
        .map((f) => ({ field: f.field, mode: "Members", includedMembers: [] })),
    }));

  // The busy/unreachable data server retries by itself after a short wait (the server usually
  // recovers within a minute); "Try now" skips the wait.
  const queryRetryIn = useRetryCountdown(Boolean(error) && errorTemporary && !loading, 20, () => runQuery(), queryFailures);
  const retryMeta = () => {
    setTablesLoading(true);
    setMetaAttempt((n) => n + 1);
  };
  const metaRetryIn = useRetryCountdown(Boolean(metaError) && metaTemporary && !tablesLoading, 20, retryMeta, metaAttempt);

  // The top bar's actions (2026-10-20, "like excel but more modern"): Export and Save Snapshot sit
  // where Excel for the web has Share — they used to be a card under the Fields pane. Report Layout and
  // Repeat All Item Labels aren't repeated in the Save popover any more: they're on the Design tab.
  const headerActions = (
    <>
      <button
        type="button"
        onClick={handleExport}
        disabled={!result || exporting}
        className="btn-secondary inline-flex items-center gap-1.5 px-3 py-1.5 text-sm disabled:cursor-not-allowed disabled:opacity-50"
        title="Download this PivotTable as an Excel file"
      >
        <FileSpreadsheet className="h-4 w-4" style={{ color: "#217346" }} />
        {/* Icon only on a phone (the title stays as its tooltip) so the top bar fits, 2026-10-25. */}
        <span className="hidden sm:inline">{exporting ? "Exporting…" : "Export"}</span>
      </button>
      <Popover
        align="right"
        panelClassName="!min-w-[20rem] !max-w-[22rem]"
        trigger={(toggle) => (
          <button
            type="button"
            onClick={toggle}
            disabled={!result}
            className="btn-primary inline-flex items-center gap-1.5 px-3 py-1.5 text-sm disabled:cursor-not-allowed disabled:opacity-50"
            title={result ? "Save this PivotTable as a snapshot" : "Build a PivotTable first"}
          >
            <Save className="h-4 w-4" />
            <span className="hidden sm:inline">Save Snapshot</span>
          </button>
        )}
      >
        {(close) => (
          <form
            className="flex flex-col gap-2"
            onSubmit={async (e) => {
              e.preventDefault();
              if (await handleSave()) close();
            }}
          >
            <p className="text-sm font-semibold text-ink">Save Snapshot</p>
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
                className="inline-flex items-center gap-1 self-start text-xs font-medium text-primary-strong hover:underline"
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

            <select
              className="w-full rounded-md border border-slate-300 bg-white px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-primary/40 focus:border-primary"
              value={tableStyle}
              onChange={(e) => setTableStyle(e.target.value)}
              title="Table style for the saved snapshot"
            >
              <option value="Default">Style: Default</option>
              <option value="Banded">Style: Banded rows</option>
              <option value="Dark">Style: Dark</option>
              <option value="Minimal">Style: Minimal</option>
            </select>

            <label className="flex items-center gap-2 text-sm text-ink">
              <input type="checkbox" checked={includeChart} onChange={(e) => setIncludeChart(e.target.checked)} />
              Include chart
            </label>
            {includeChart && (
              <select
                className="w-full rounded-md border border-slate-300 bg-white px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-primary/40 focus:border-primary"
                value={chartType}
                onChange={(e) => setChartType(e.target.value)}
                title="Chart type"
              >
                <option value="Bar">Chart: Bar</option>
                <option value="Line">Chart: Line</option>
                <option value="Pie">Chart: Pie</option>
              </select>
            )}

            <button type="submit" disabled={saving || !snapshotName.trim()} className="btn-primary mt-1 disabled:cursor-not-allowed disabled:opacity-50">
              {saving ? "Saving…" : "Save"}
            </button>
          </form>
        )}
      </Popover>
    </>
  );

  return (
    // onDragOver/onDrop here, outside AppShell entirely: catches a field dropped ANYWHERE on the
    // page (including the sidebar, the grid, blank margins) that no more specific zone inside
    // FieldPicker already claimed — see handlePageDrop's own WHY comment.
    <div onDragOver={handlePageDragOver} onDrop={handlePageDrop}>
      <AppShell workspace actions={headerActions}>
        {metaError &&
          (metaTemporary ? (
            <MessageBar tone="info" action={{ label: "Try now", onClick: retryMeta, disabled: tablesLoading }}>
              {metaError} The field list will load as soon as it answers{metaRetryIn ? ` — trying again in ${metaRetryIn}s` : "…"}
            </MessageBar>
          ) : (
            <MessageBar tone="warning">Couldn&apos;t load the field list: {metaError}</MessageBar>
          ))}

      {/* Excel's own window: ribbon + worksheet on the left, the PivotTable Fields pane docked on
          the right (2026-10-04, "زي الإكسل"), the status bar along the bottom.
          Workspace layout (2026-10-20 UI refresh): from xl (1280px) up the page is exactly the
          viewport's height (AppShell `workspace`) — this row takes what the status bar leaves, the
          grid flexes into the worksheet column's leftover height and scrolls inside itself, and the
          Fields pane scrolls on its own and is resizable (PaneSplitter) and collapsible to a rail.
          That replaces the fixed 460px sticky pane beside a grid that hugged its table, which left a
          big empty band on wide screens. Below xl the pane drops UNDER the grid at full width
          (2026-10-19, "it's not responsive") and the page scrolls normally. */}
      <div className="flex flex-col gap-3 lg:min-h-0 lg:flex-1 lg:flex-row lg:gap-0">
        {/* min-w-0: without it, a wide pivot grid can force this whole row (and the page) wider
            than the viewport — the classic flexbox min-width:auto trap. */}
        <div className="flex min-w-0 flex-col gap-2 lg:min-h-0 lg:flex-1">
          {/* No "Run query" button — Excel re-queries automatically (2026-10-04); the status bar's
              "Updating…" is the only trace of a recalculation, like Excel's. */}
          {error &&
            (errorTemporary ? (
              <MessageBar tone="info" action={{ label: "Try now", onClick: () => runQuery(), disabled: loading }}>
                {error} Your PivotTable will appear as soon as it answers{queryRetryIn ? ` — trying again in ${queryRetryIn}s` : "…"}
              </MessageBar>
            ) : (
              <MessageBar tone="error" action={{ label: "Retry", onClick: () => runQuery(), disabled: loading }}>
                {error}
              </MessageBar>
            ))}

          {/* The ribbon sits flush at the column's top (2026-10-12, "الجزء دا عايز يطلع لفوق").
              This wrapper is itself a flex column so PivotGrid's `fill` box (its last child) takes
              the leftover height. */}
          <div className="flex min-h-0 flex-1 flex-col gap-2">
            <PivotToolbar
              request={editingRequest}
              onChange={editRequest}
              measureOptions={measures.map((m) => ({ value: m.field, label: m.displayName }))}
              dimensions={dimensions}
              allDimensions={allDimensions}
              pivotOptions={pivotOptions}
              onPivotOptionsChange={setPivotOptions}
              onOpenPivotTableOptions={() => setShowPivotOptions(true)}
              fieldListShown={showFieldList}
              onToggleFieldList={() => setShowFieldList((v) => !v)}
              layout={{
                reportLayout,
                repeatItemLabels,
                fieldLayouts,
                onReportLayoutChange: setReportLayout,
                onRepeatItemLabelsChange: setRepeatItemLabels,
                onFieldLayoutsChange: setFieldLayouts,
              }}
              // Home › Number (2026-10-20): like Excel, formats just the selected cells (the grid
              // applies it — see cellFormats). Display-only: no re-query.
              history={history}
              numberFormat={{
                selection: selectionSummary,
                onApply: (fn) => gridCommandRef.current?.applyNumberFormat(fn),
              }}
              actions={{
                onClearAll: handleClearAll,
                onClearFilters: handleClearFilters,
                // Report Filter fields = the Filters area's fields (one still at "(All)" too — it isn't
                // in the query, but it's exactly the field Excel offers), not on Rows/Columns, and not
                // a slicer's/timeline's (they aren't in the Report Filter area).
                reportFilterFields: lastQuery ? reportFilterFieldsOf(request).map((field) => ({ field, label: fieldLabels([field], dimensions)[0] })) : [],
                onShowReportFilterPages: handleShowReportFilterPages,
                filterPagesStatus,
              }}
            />
            {showPivotOptions && (
              <PivotTableOptionsDialog
                options={pivotOptions}
                request={editingRequest}
                onApply={(options, requestPatch) => {
                  setPivotOptions(options);
                  // Only touch `request` when a query-affecting setting really changed (it re-runs).
                  // Turning "Allow multiple filters per field" off keeps one filter per field (the
                  // backend would reject the rest).
                  const differs = (key, value) => (typeof value === "boolean" ? Boolean(editingRequest[key]) !== value : JSON.stringify(editingRequest[key]) !== JSON.stringify(value));
                  if (Object.entries(requestPatch).some(([key, value]) => differs(key, value))) {
                    editRequest({
                      ...editingRequest,
                      ...requestPatch,
                      filters: requestPatch.allowMultipleFiltersPerField ? editingRequest.filters : onePerField(editingRequest.filters),
                    });
                  }
                }}
                onClose={() => setShowPivotOptions(false)}
              />
            )}
            <PivotGrid
              result={result}
              rows={lastQuery?.rows}
              columns={lastQuery?.columns}
              rowFieldLabels={fieldLabels(lastQuery?.rows, dimensions)}
              columnFieldLabels={fieldLabels(hideMeasuresLabelIfSingleValue(lastQuery?.columns, lastQuery?.values?.length ?? 0), dimensions, fieldLayouts)}
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
                  editRequest({ ...editingRequest, rowsHierarchyDrilldown: { ...current, scopedExpansions: nextScoped } });
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
                editRequest({ ...editingRequest, rowsHierarchyDrilldown: { ...current, expandedMemberUniqueNames: expanded, expandedAncestors: ancestors } });
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
                editRequest({ ...editingRequest, columnsHierarchyDrilldown: toggleDrilldownExpansion(current, uniqueName, outer, ancestors) });
              }}
              // request/onChange, not lastQuery: a header's Sort/Filter menu edits the LIVE field
              // config (same request/setRequest pair FieldPicker already uses) — rows/columns above
              // stay keyed off lastQuery so the field identifiers always match what's actually
              // displayed, but the edit itself targets the live request, same as every other
              // Sort/Filter control in FieldPicker (still needs Run query to take effect — this isn't
              // a special case, just a new entry point into the same existing flow).
              request={editingRequest}
              onChange={editRequest}
              onRefresh={runQuery}
              onHideFieldList={() => setShowFieldList((v) => !v)}
              fieldListShown={showFieldList}
              pivotOptions={pivotOptions}
              onOpenPivotTableOptions={() => setShowPivotOptions(true)}
              fill
              zoom={zoom}
              onSelectionStatsChange={setSelectionSummary}
              cellFormats={cellFormats}
              onCellFormatsChange={setCellFormats}
              commandRef={gridCommandRef}
              loading={loading}
            />
          </div>

          {/* The chart takes its natural height under the grid, capped (and scrolling) so it can't
              squeeze the grid away in the fixed-height workspace. */}
          {includeChart && (
            <div className="custom-scrollbar shrink-0 lg:max-h-[45%] lg:overflow-auto">
              {chartUnavailableReason(result, lastQuery) && (
                <p className="rounded-md border border-slate-200 bg-slate-50 px-3 py-2 text-sm text-muted">{chartUnavailableReason(result, lastQuery)}</p>
              )}
              <PivotChart result={result} valueFields={displayValues} chartType={chartType} rows={lastQuery?.rows} />
            </div>
          )}
        </div>

        {/* The Fields pane. Default 460px (2026-10-04: "المفروض ان المساحة الكبرى تبقى للداتا" — the
            worksheet dominates, like Excel's narrow fixed pane), now draggable 340–760px. Its
            content is meant to FIT without a scrollbar (2026-10-12, "انا ماقولتلكش زود scroll
            bar!!"); overflow-y-auto below is only the safety net for a short window, never shown
            when it fits. */}
        {showFieldList ? (
          <>
          <PaneSplitter
            width={fieldPaneWidth}
            min={FIELD_PANE_MIN_WIDTH}
            max={FIELD_PANE_MAX_WIDTH}
            defaultWidth={FIELD_PANE_DEFAULT_WIDTH}
            onChange={setFieldPaneWidth}
            label="Resize the PivotTable Fields pane"
          />
          {/* Its own vertical scroll from xl up (the workspace is viewport-high, so a tall pane has
              to scroll somewhere — and never the page). Popovers inside portal to <body>, so this
              overflow can't clip them (the 2026-10-12 reason this pane once had no scroll box). */}
          <aside
            aria-label="PivotTable Fields"
            style={{ "--field-pane-width": `${fieldPaneWidth}px` }}
            className="flex w-full shrink-0 flex-col rounded-lg border border-slate-200 bg-white shadow-sm lg:min-h-0 lg:w-[min(var(--field-pane-width),40vw)] xl:w-[var(--field-pane-width)]"
          >
            {/* Excel's task pane header: its title and a close button (= Hide Field List). */}
            <div className="flex h-10 shrink-0 items-center justify-between border-b border-slate-200 pl-3 pr-1.5">
              <h2 className="text-sm font-semibold text-ink">PivotTable Fields</h2>
              <div className="ml-auto flex items-center" ref={setPaneToolsEl} />
              <button
                type="button"
                onClick={() => setShowFieldList(false)}
                className="rounded p-1 text-slate-400 hover:bg-slate-100 hover:text-ink"
                aria-label="Close the PivotTable Fields pane"
                title="Close"
              >
                <X className="h-4 w-4" />
              </button>
            </div>
            <div className="custom-scrollbar min-h-0 flex-1 overflow-x-hidden p-2 lg:overflow-y-auto">
            <FieldPicker
              value={pendingRequest ?? request}
              onChange={deferLayout ? (next) => setPendingRequest(normalizeMeasuresPlacement(next)) : setRequest}
              dimensions={dimensions}
              measures={measures}
              tables={tables}
              tablesLoading={tablesLoading}
              enabledTableNames={enabledTableNames}
              onEnabledTableNamesChange={updateEnabledTables}
              fieldsLoading={fieldsLoading}
              allDimensions={allDimensions}
              fieldLayouts={fieldLayouts}
              reportLayout={reportLayout}
              repeatItemLabels={repeatItemLabels}
              onFieldLayoutsChange={setFieldLayouts}
              toolsTarget={paneToolsEl}
            />
            </div>
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
            {/* Excel's "Defer Layout Update" + "Update", at the bottom of the Field List. */}
            <div className="flex shrink-0 items-center justify-between gap-2 border-t border-slate-200 px-3 py-2 text-sm">
              <label className="flex items-center gap-2 text-ink">
                <input
                  type="checkbox"
                  className="h-4 w-4 accent-primary"
                  checked={deferLayout}
                  onChange={(e) => {
                    setDeferLayout(e.target.checked);
                    if (!e.target.checked) applyPendingLayout();
                  }}
                />
                Defer Layout Update
              </label>
              <button
                type="button"
                onClick={applyPendingLayout}
                disabled={!deferLayout || !pendingRequest}
                className="btn-secondary px-3 py-1 text-xs disabled:cursor-not-allowed disabled:opacity-50"
              >
                Update
              </button>
            </div>
          </aside>
          </>
        ) : (
          // Excel's "Hide Field List" (ribbon toggle, value-cell menu, status bar) collapses the pane
          // to this slim rail along the right edge (2026-10-20) — one click brings it back.
          <button
            type="button"
            onClick={() => setShowFieldList(true)}
            className="flex shrink-0 items-center justify-center gap-2 rounded-md border border-slate-200 bg-white px-3 py-2 text-xs font-medium text-slate-600 hover:border-primary hover:text-ink lg:ml-2 lg:w-8 lg:flex-col lg:px-0 lg:py-3"
            title="Show the PivotTable Fields pane"
          >
            <PanelRightOpen className="h-4 w-4 shrink-0" />
            <span className="lg:[writing-mode:vertical-rl]">PivotTable Fields</span>
          </button>
        )}
      </div>
      <Toasts
        items={[
          saveMessage && { tone: "success", text: saveMessage, onClose: () => setSaveMessage(null) },
          saveError && { tone: "error", text: saveError, onClose: () => setSaveError(null) },
          exportError && { tone: "error", text: exportError, onClose: () => setExportError(null) },
          exportNotice && { tone: "warning", text: exportNotice, onClose: () => setExportNotice(null) },
        ].filter(Boolean)}
      />
      <StatusBar
        loading={loading}
        onCancel={handleCancelQuery}
        pendingLayout={deferLayout && Boolean(pendingRequest)}
        stats={selectionSummary}
        zoom={zoom}
        onZoomChange={setZoom}
        fieldListShown={showFieldList}
        onToggleFieldList={() => setShowFieldList((v) => !v)}
      />
      </AppShell>
    </div>
  );
}
