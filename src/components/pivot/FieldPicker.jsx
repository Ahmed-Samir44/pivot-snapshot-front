import { useEffect, useRef, useState } from "react";
import Select from "react-select";
import { X, GripVertical, Plus, Filter, Columns3, Rows3, Sigma, Loader2, ChevronDown, ChevronRight, ChevronLeft, Search, Folder, Table2 } from "lucide-react";
import HierarchicalCubePathSelect from "../forms/HierarchicalCubePathSelect";
import MultiSelectField from "../forms/MultiSelectField";
import MemberCheckboxList from "../forms/MemberCheckboxList";
import Popover from "../forms/Popover";
import { getMembers } from "../../services/cubeMetaApi";
import { formatNumber } from "../../utils/numberFormat";
import { formatDimensionName } from "../../utils/cubeMeta";
import {
  setFieldDragPayload,
  readFieldDragPayload,
  removeFieldEverywhere,
  clearOrphanedCalculatedItems,
  clearDanglingMeasureReferences,
  clearDanglingBaseFieldReferences,
  reconcileDrilldowns,
} from "../../utils/dragDrop";
import { MEASURES_PLACEHOLDER } from "../../utils/measuresPlacement";
import { DEFAULT_MIN_COLOR, DEFAULT_MID_COLOR, DEFAULT_MAX_COLOR, DEFAULT_BAR_COLOR } from "../../utils/conditionalFormat";

export const SHOW_VALUES_AS_OPTIONS = [
  { value: "Normal", label: "Normal" },
  { value: "PercentOfGrandTotal", label: "% of Grand Total" },
  { value: "RunningTotal", label: "Running Total" },
  { value: "Rank", label: "Rank Largest to Smallest" },
  { value: "PercentOfParentRow", label: "% of Parent Row" },
  { value: "DifferenceFrom", label: "Difference From (previous)" },
  { value: "PercentOfRowTotal", label: "% of Row Total" },
  { value: "PercentOfColumnTotal", label: "% of Column Total" },
  { value: "PercentOfParentColumn", label: "% of Parent Column" },
  { value: "Index", label: "Index" },
  { value: "PercentDifferenceFrom", label: "% Difference From (previous)" },
  { value: "PercentRunningTotal", label: "% Running Total In" },
  { value: "RankAscending", label: "Rank Smallest to Largest" },
  { value: "PercentOf", label: "% Of…" },
];

// Every Show Values As variant whose output is a ratio meant to be read as a percentage (see
// ShowValuesAsProcessor.cs's own class-level WHY comment — they all produce 0.3843, not 38.43,
// since 2026-10-05). Used below to auto-switch Number Format to Percentage the moment one of these
// is picked, matching real Excel's own behavior (Excel auto-applies a percentage format when you
// choose a % Show Values As option) — without this, the default "General" format would show the
// raw ratio (e.g. "0.38") instead of a readable percentage.
export const PERCENT_SHOW_VALUES_AS = new Set([
  "PercentOfGrandTotal",
  "PercentOfRowTotal",
  "PercentOfColumnTotal",
  "PercentOfParentRow",
  "PercentOfParentColumn",
  "PercentDifferenceFrom",
  "PercentRunningTotal",
  "PercentOf",
]);

// currencySymbol: "" (not any specific currency) — which symbol to show, if any, is a per-report
// choice only the user makes. An earlier default of "EGP" was hardcoded here without ever having
// been asked for; flagged live, 2026-10-06 ("ليه محدد ال currency egp ليه بتاخد قرارات من عندك").
export const DEFAULT_FORMAT = { type: "General", decimalPlaces: 2, currencySymbol: "" };

const nativeSelectClass =
  "rounded-xl border border-slate-200 bg-white/90 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent";


// Excel's Field List: every cube DIMENSION field in one place, dragged from here into whichever
// of Filters/Rows/Columns it belongs in. Unlike the zones, a field never "leaves" this list when
// placed somewhere — Excel's own field list keeps every field checkbox visible too — it just gets
// a small dot marking it as already used somewhere. A field can still be placed in Values AT THE
// SAME TIME as Rows/Columns/Filters (that's a different kind of area — see moveField's own WHY
// comment) but, since 2026-10-05, NOT in more than one of Filters/Rows/Columns at once — matching
// real Excel's own rule exactly.
//
// Measures ARE listed here too (2026-10-05, matching real Excel's own Field List exactly — see
// buildFolderTree's own WHY comment), draggable the same way as any dimension field. The original
// concern that kept them out entirely — a real bug, 2026-09-23, where a mis-dropped measure in
// Rows/Columns produced broken MDX — is handled at the DROP TARGET instead now: every
// isFieldAccepted check on the Rows/Columns/Filters zones below only accepts real dimensionOptions
// (plus the Σ Values sentinel on Rows/Columns), so a measure dropped there is simply rejected,
// regardless of where it was dragged from. Values' own multi-select above the pill list is still
// the primary way to pick measures; dragging one here is just an equally-valid alternative.
//
// Grouped by dimension (the text before the first "." in displayName, e.g. "Patient" for
// "Patient.Patient Name") rather than one flat wrapped wall of pills — a real cube has 20+ fields
// across 8+ dimensions, and a single ungrouped list of same-weight buttons became unreadable at
// that size (caught live 2026-09-22: "ايه الشكل القذر دا"). Matches Excel's own Field List, which
// is always grouped by table/dimension with a bold header per group, never one flat list.
//
// A Values entry's own effective caption — what actually renders as its column/row header — is its
// customLabel if set, otherwise (mirrors ValuePill's own computedLabel, and
// MdxPivotQueryBuilder.BuildMeasureExpression's naming on the backend) a dimension field's
// "<Aggregation> of <field>" or a real measure's own plain name.
function effectiveValueCaption(v) {
  if (v.customLabel) return v.customLabel;
  if (!v.field.startsWith("[")) return v.field;
  return `${v.dimensionAggregation ?? "Count"} of ${formatDimensionName(v.field)}`;
}

// See moveField's own WHY comment on why every Values entry needs a unique caption — real Excel's
// own "Sum of Revenue2"-style naming for a duplicated field, just appending a number instead of
// Excel's own concatenated-no-space style for a cube-measure field whose name doesn't start with a
// recognizable word Excel would glue a suffix onto so awkwardly.
function nextAvailableDuplicateLabel(field, existingValues) {
  const baseLabel = field.startsWith("[") ? formatDimensionName(field) : field;
  const used = new Set(existingValues.map(effectiveValueCaption));
  let n = 2;
  while (used.has(`${baseLabel} ${n}`)) {
    n += 1;
  }
  return `${baseLabel} ${n}`;
}

// Builds a nested folder tree from each field's own `displayFolder` (backslash-delimited, straight
// from the cube's own MEASURE_DISPLAY_FOLDER/HIERARCHY_DISPLAY_FOLDER — see the backend's CubeField
// .DisplayFolder WHY comment) — matches real Excel's own Field List exactly: a table's fields
// aren't always one flat list, some are nested a folder (or several) deep (confirmed live
// 2026-10-05, e.g. "Fact Billing" groups dozens of measures into "Historical Measures", "Power
// BI\OPD New Dashboard", etc.). A field with no displayFolder sits directly in its table/folder's
// own `fields` array, same depth as the folders sitting beside it.
function buildFolderTree(fields) {
  const root = { fields: [], folders: new Map() };
  for (const f of fields) {
    let node = root;
    if (f.displayFolder) {
      for (const part of f.displayFolder.split("\\").map((p) => p.trim()).filter(Boolean)) {
        if (!node.folders.has(part)) node.folders.set(part, { fields: [], folders: new Map() });
        node = node.folders.get(part);
      }
    }
    node.fields.push(f);
  }
  return root;
}

// Groups a folder-tree level's own flat field list by shared `hierarchyUniqueName` — every level
// of a multi-level natural hierarchy (e.g. "Specialty Group"/"Specialty"/"Doctor Name", all under
// "Specialty Hierarchy") lands in the SAME node.fields array today (they all inherit the same
// displayFolder from their hierarchy), so grouping them here is enough to offer a single "+ Add
// whole hierarchy" action instead of the user hunting for and adding each level separately (OLAP
// +/- drill-down Phase 1, 2026-10-07 — see EXCEL_PARITY_AUDIT.md). A "hierarchy" with only one
// exposed level (or no hierarchyUniqueName at all, e.g. measures) renders as a plain field row —
// the whole-hierarchy action only makes sense with 2+ levels to add.
function groupFieldsByHierarchy(fields) {
  const groups = [];
  const byHierarchy = new Map();
  for (const f of fields) {
    if (!f.hierarchyUniqueName) {
      groups.push({ type: "field", field: f });
      continue;
    }
    let group = byHierarchy.get(f.hierarchyUniqueName);
    if (!group) {
      group = { type: "hierarchy", hierarchyUniqueName: f.hierarchyUniqueName, fields: [] };
      byHierarchy.set(f.hierarchyUniqueName, group);
      groups.push(group);
    }
    group.fields.push(f);
  }
  return groups.map((g) =>
    g.type === "hierarchy" && g.fields.length < 2 ? { type: "field", field: g.fields[0] } : g,
  );
}

// Strips the leading "<Table>." real Excel-formatted dimension labels already drop (see
// dimensionFieldsByTable's own `label` computation below) — same convention, applied to a
// hierarchy's own unique name instead of a single level's.
function hierarchyGroupLabel(hierarchyUniqueName) {
  const formatted = formatDimensionName(hierarchyUniqueName);
  const dotIndex = formatted.indexOf(".");
  return dotIndex === -1 ? formatted : formatted.slice(dotIndex + 1);
}

// Renders one folder-tree level: this level's own fields first (draggable AND checkbox-able,
// indented by `depth`), then its sub-folders (each its own collapsed-by-default expand/collapse
// row), recursing one level deeper per nested folder — mirrors real Excel's own indentation for a
// multi-level Display Folder.
function FolderTreeNode({ node, folderKeyPrefix, depth, expandedFolders, toggleFolder, isPlaced, onFieldCheckToggle, onAddHierarchy, onStartDrilldown }) {
  const indent = 28 + depth * 16;
  const renderFieldRow = ({ field, label }, rowIndent) => (
    <li
      key={field}
      draggable
      onDragStart={(e) => setFieldDragPayload(e, field, "available")}
      style={{ paddingLeft: rowIndent }}
      className="flex cursor-grab items-center gap-1.5 py-1 pr-2 font-mono text-[11.5px] text-ink hover:bg-slate-50 active:cursor-grabbing"
    >
      {/* Real Excel's own click-to-add: checking a field's box adds it straight to its default
          area (a measure to Values, a dimension attribute to Rows) without needing to drag it —
          confirmed live 2026-10-05 against a real Excel screenshot, where every leaf field has
          its own checkbox. Unchecking removes it from WHEREVER it currently sits (not just the
          default area — a field dragged somewhere else still unchecks cleanly). */}
      <input
        type="checkbox"
        checked={isPlaced(field)}
        onChange={(e) => onFieldCheckToggle(field, e.target.checked)}
        onClick={(e) => e.stopPropagation()}
        className="h-3 w-3 shrink-0 accent-gold"
      />
      <span className="truncate">{label}</span>
    </li>
  );
  return (
    <>
      {groupFieldsByHierarchy(node.fields).map((group) => {
        if (group.type === "field") {
          return renderFieldRow(group.field, indent);
        }
        const ordered = [...group.fields].sort((a, b) => (a.levelNumber ?? 0) - (b.levelNumber ?? 0));
        const fieldsInOrder = ordered.map((f) => f.field);
        return (
          <li key={group.hierarchyUniqueName}>
            <div className="flex items-center gap-1.5 py-1 pr-2" style={{ paddingLeft: indent }}>
              <span className="truncate text-[11px] font-semibold text-muted" title={group.hierarchyUniqueName}>
                {hierarchyGroupLabel(group.hierarchyUniqueName)}
              </span>
              {onAddHierarchy && (
                <span className="ml-auto flex shrink-0 gap-1">
                  <button
                    type="button"
                    onClick={() => onAddHierarchy(fieldsInOrder, "rows")}
                    className="rounded border border-slate-200 px-1.5 py-0.5 text-[10px] text-muted hover:bg-slate-50 hover:text-ink"
                    title={`Add all ${fieldsInOrder.length} levels of this hierarchy to Rows, in order`}
                  >
                    + Rows
                  </button>
                  <button
                    type="button"
                    onClick={() => onAddHierarchy(fieldsInOrder, "columns")}
                    className="rounded border border-slate-200 px-1.5 py-0.5 text-[10px] text-muted hover:bg-slate-50 hover:text-ink"
                    title={`Add all ${fieldsInOrder.length} levels of this hierarchy to Columns, in order`}
                  >
                    + Cols
                  </button>
                </span>
              )}
              {/* Real Excel's own per-member +/- drilldown (PLAN 8, 2026-10-13, EXPERIMENTAL — see
                  plan file WHY comment) — this button's own label/placement is a plain functional
                  entry point, NOT a reproduction of a specific real-Excel dialog's exact wording or
                  look (unlike every other dialog built this session), since no screenshot exists yet
                  for how Excel's own field list visually distinguishes a drillable hierarchy field.
                  Deliberately next to "+ Rows"/"+ Cols", not replacing them — v1 Rows-only. */}
              {onStartDrilldown && (
                <button
                  type="button"
                  onClick={() => onStartDrilldown(fieldsInOrder[0])}
                  className="ml-1 shrink-0 rounded border border-slate-200 px-1.5 py-0.5 text-[10px] text-muted hover:bg-slate-50 hover:text-ink"
                  title="Add this hierarchy to Rows with real per-member +/- drilldown (experimental)"
                >
                  Drilldown
                </button>
              )}
            </div>
            <ul>{ordered.map((f) => renderFieldRow(f, indent + 16))}</ul>
          </li>
        );
      })}
      {Array.from(node.folders.keys())
        .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }))
        .map((folderName) => {
          const key = `${folderKeyPrefix}\u0001${folderName}`;
          const isOpen = expandedFolders.has(key);
          const child = node.folders.get(folderName);
          return (
            <li key={key}>
              <button
                type="button"
                onClick={() => toggleFolder(key)}
                style={{ paddingLeft: indent }}
                className="flex w-full items-center gap-1.5 py-1 pr-2 text-left hover:bg-slate-50"
              >
                {isOpen ? (
                  <ChevronDown className="h-3 w-3 shrink-0 text-muted" />
                ) : (
                  <ChevronRight className="h-3 w-3 shrink-0 text-muted" />
                )}
                <Folder className="h-3 w-3 shrink-0 text-gold" />
                <span className="truncate text-xs text-ink">{folderName}</span>
              </button>
              {isOpen && (
                <ul>
                  <FolderTreeNode
                    node={child}
                    folderKeyPrefix={key}
                    depth={depth + 1}
                    expandedFolders={expandedFolders}
                    toggleFolder={toggleFolder}
                    isPlaced={isPlaced}
                    onFieldCheckToggle={onFieldCheckToggle}
                    onAddHierarchy={onAddHierarchy}
                    onStartDrilldown={onStartDrilldown}
                  />
                </ul>
              )}
            </li>
          );
        })}
    </>
  );
}

function AvailableFieldsPanel({ dimensions, measures, valuesCount, isPlaced, tables, tablesLoading, enabledTableNames, onEnabledTableNamesChange, fieldsLoading, onFieldDropped, onAddHierarchy, onStartDrilldown }) {
  // Real Excel's own Field List shows EVERY table immediately, collapsed, with its own ▸/▾ arrow —
  // no prior "search and pick a table" step at all (confirmed live, 2026-10-05, against real
  // Excel's own PivotTable Fields pane screenshots: ~300 tables, all listed, none pre-selected).
  // The earlier "search and pick" MultiSelectField UI was a deliberate scope decision to dodge
  // fetching 2301 fields up front — but it was also a VISIBLE deviation from Excel's actual field
  // list shape, not just an internal implementation detail, and was flagged as a real difference
  // once compared side-by-side with the real dialog. This keeps the SAME lazy-fetch-per-table
  // contract (enabledTableNames/onEnabledTableNamesChange — PivotBuilder.jsx's own effect that
  // fetches a table's columns only once it's "enabled" needs no change at all) but now "enabled"
  // means "expanded in the tree", exactly matching Excel's own expand-to-load interaction, with a
  // search box that FILTERS the visible table list (like Excel's own) rather than gating it.
  const [tableSearch, setTableSearch] = useState("");
  // Which "table>>folder\path" keys are expanded — purely a display nicety (every field for an
  // expanded table is already fetched/available, folders don't gate anything), so plain local
  // state is enough; collapsed by default, same as a table itself.
  const [expandedFolders, setExpandedFolders] = useState(() => new Set());
  const toggleFolder = (key) => {
    setExpandedFolders((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };
  // Collapses every expanded table AND every expanded folder in one click — requested live,
  // 2026-10-05. Doesn't touch which fields are actually placed on the pivot, just the tree's own
  // open/closed state (collapsing an already-fetched table keeps its columns cached, so re-
  // expanding it afterward is instant, same as collapsing it one at a time already was).
  const collapseAll = () => {
    onEnabledTableNamesChange([]);
    setExpandedFolders(new Set());
  };
  // Checking a field adds it to its default area (a measure — a plain name, no "[" prefix — to
  // Values; a dimension attribute to Rows), matching real Excel's own click-to-add default.
  // Unchecking removes it from wherever it currently sits, not just that default area — reuses
  // moveField's own toZone==="available" branch, which already ignores fromZone and removes
  // everywhere (see moveField's own WHY comment), so any non-"available" placeholder works here.
  const onFieldCheckToggle = (field, checked) => {
    if (!checked) {
      onFieldDropped(field, "checkbox", "available");
      return;
    }
    onFieldDropped(field, "available", field.startsWith("[") ? "rows" : "values");
  };
  // Dragging an already-placed field back here removes it from wherever it is — Excel's own "drag
  // a field out of any area to remove it", as an addition alongside each pill's existing X button
  // (requested live, 2026-09-28). This panel never ACCEPTS a field as a new placement (dragging
  // FROM here only adds elsewhere), so every drop it sees is necessarily a removal.
  const [dragOverRemove, setDragOverRemove] = useState(false);
  const handleRemoveDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOverRemove(false);
    const payload = readFieldDragPayload(e);
    if (payload && payload.source !== "available" && onFieldDropped) {
      onFieldDropped(payload.field, payload.source, "available");
    }
  };
  // Real Excel does NOT lump every measure under one generic "Σ Measures" bucket (confirmed live
  // 2026-10-05, against a real Excel screenshot) — each measure group (e.g. "Fact Billing") is its
  // OWN top-level Field List entry, same as a dimension table, with that measure group's own
  // measures inside it (further organized into DisplayFolder sub-folders — see the backend's
  // CubeField.MeasureGroupName/DisplayFolder WHY comments). A measure group can ALSO be a real,
  // separately-fetchable dimension table (this cube exposes "Fact Billing" as a browsable "fact
  // dimension" too) — in that case its dimension attributes and its measures share the exact same
  // tree entry, exactly like Excel shows them.
  const measuresByGroup = new Map();
  measures.forEach((m) => {
    const groupName = m.measureGroupName || "Measures";
    if (!measuresByGroup.has(groupName)) measuresByGroup.set(groupName, []);
    measuresByGroup.get(groupName).push({ field: m.field, label: m.displayName, displayFolder: m.displayFolder ?? null });
  });

  const dimensionFieldsByTable = new Map();
  // `dimensions` only ever holds columns for tables PivotBuilder has already fetched (i.e. every
  // currently-EXPANDED table — enabledTableNames now means "expanded", see this component's own
  // top-of-function WHY comment) — so a table not yet in this map just hasn't loaded its columns
  // yet (or has none).
  dimensions.forEach((d) => {
    const dotIndex = d.displayName.indexOf(".");
    const groupName = dotIndex === -1 ? d.displayName : d.displayName.slice(0, dotIndex);
    const label = dotIndex === -1 ? d.displayName : d.displayName.slice(dotIndex + 1);
    if (!dimensionFieldsByTable.has(groupName)) dimensionFieldsByTable.set(groupName, []);
    dimensionFieldsByTable.get(groupName).push({
      field: d.field,
      label,
      displayFolder: d.displayFolder ?? null,
      hierarchyUniqueName: d.hierarchyUniqueName ?? null,
      levelNumber: d.levelNumber ?? null,
    });
  });

  const fieldsFor = (name) => [...(dimensionFieldsByTable.get(name) ?? []), ...(measuresByGroup.get(name) ?? [])];

  const toggleTable = (name) => {
    onEnabledTableNamesChange(
      enabledTableNames.includes(name) ? enabledTableNames.filter((n) => n !== name) : [...enabledTableNames, name],
    );
  };

  // Union of every real cube table AND every measure group name — a measure group with no matching
  // dimension table still needs its own top-level entry (most of this cube's measure groups aren't
  // ALSO set up as a browsable fact dimension). Deduplicated since a name can appear in both.
  const allTableNames = Array.from(new Set([...tables.map((t) => t.displayName), ...measuresByGroup.keys()])).sort(
    (a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }),
  );
  const query = tableSearch.trim().toLowerCase();
  const visibleTableNames = query ? allTableNames.filter((name) => name.toLowerCase().includes(query)) : allTableNames;

  // Real Excel's own "Σ Values" pseudo-field — a single field (not a table with children),
  // draggable into Rows or Columns to choose where the values axis appears. Only listed once there
  // are 2+ measures in Values: with fewer, there's no position to choose (Excel doesn't offer it
  // as an item either in that case).
  const showValuesPseudoField = valuesCount >= 2 && (!query || "σ values".includes(query) || "values".includes(query));

  return (
    // h-full flex flex-col: fills the exact pixel height its caller's wrapper div sets (from
    // zonesColumnHeight — see FieldPicker's own WHY comment) instead of sizing to its own content;
    // only the scrollable list below (flex-1 min-h-0) actually grows/shrinks into whatever's left
    // after the header and search box above, which stay their natural height either way.
    <div className="card flex h-full flex-col">
      <div className="mb-1.5 flex items-center justify-between">
        <p className="text-xs font-bold uppercase tracking-wide text-muted">Fields</p>
        <button type="button" onClick={collapseAll} className="btn-secondary px-2 py-1 text-xs">
          Collapse All
        </button>
      </div>
      {/* Real Excel's own Field List: EVERY table listed immediately (no prior pick step), each
          collapsed behind its own ▸ arrow — a table's columns are only fetched once ITS arrow is
          clicked (same lazy-fetch PivotBuilder.jsx already did per "enabled" table, just triggered
          by expand/collapse now instead of a multiselect pick). The search box FILTERS which table
          rows are visible, exactly like Excel's own — it never gates fetching. */}
      {tablesLoading ? (
        <p className="mb-2 flex items-center gap-1.5 text-xs text-muted">
          <Loader2 className="h-3.5 w-3.5 animate-spin" />
          Loading the list of tables…
        </p>
      ) : (
        <div className="relative mb-2">
          <Search className="pointer-events-none absolute left-2.5 top-1/2 h-3.5 w-3.5 -translate-y-1/2 text-muted" />
          <input
            type="text"
            value={tableSearch}
            onChange={(e) => setTableSearch(e.target.value)}
            placeholder="Search"
            className="w-full rounded-md border border-slate-200 bg-white py-1.5 pl-8 pr-3 text-xs focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent"
          />
        </div>
      )}
      <div
        // flex-1 min-h-0: fills whatever's left of the card's own height (set by its caller from
        // the zones column's measured height) after the header/search box above. min-h-0 overrides
        // a flex child's default min-height:auto (its content's size), which would otherwise keep
        // forcing this box tall enough to show everything un-scrolled no matter what flex-1 says —
        // with it, overflow-y-auto actually scrolls within whatever height it's given, same as a
        // fixed max-height would, just dynamically sized to match the zones column instead.
        className={`min-h-0 flex-1 overflow-y-auto rounded-lg border border-slate-200 ${
          dragOverRemove ? "bg-red-50 ring-2 ring-red-300 ring-inset" : ""
        }`}
        onDragOver={(e) => {
          e.preventDefault();
          setDragOverRemove(true);
        }}
        onDragLeave={() => setDragOverRemove(false)}
        onDrop={handleRemoveDrop}
      >
        {visibleTableNames.map((name, tableIndex) => {
          const isExpanded = enabledTableNames.includes(name);
          const fields = fieldsFor(name);
          // A table only has a dimension-fetch in flight if it's a REAL dimension table (present in
          // `tables`) — a pure measure group (no matching fact-dimension) has nothing to fetch, its
          // measures are already loaded eagerly, so it's never "loading". fieldsLoading itself is
          // one flag covering every table currently being fetched (see PivotBuilder.jsx) — good
          // enough to show "Loading…" under whichever real table(s) don't have columns back yet,
          // without tracking per-table in-flight state separately.
          const isRealTable = tables.some((t) => t.displayName === name);
          const isLoadingThisTable = isExpanded && fields.length === 0 && isRealTable && fieldsLoading;
          const tree = buildFolderTree(fields);
          return (
            <div key={name} className={tableIndex > 0 ? "border-t border-slate-100" : ""}>
              <button
                type="button"
                onClick={() => toggleTable(name)}
                className="flex w-full items-center gap-1.5 bg-white px-2 py-1.5 text-left hover:bg-slate-50"
              >
                {isExpanded ? (
                  <ChevronDown className="h-3.5 w-3.5 shrink-0 text-muted" />
                ) : (
                  <ChevronRight className="h-3.5 w-3.5 shrink-0 text-muted" />
                )}
                {/* Σ for a measure group, a plain table icon otherwise — matches real Excel's own
                    Field List exactly (confirmed live 2026-10-05): a table that has measures gets
                    the Σ glyph even when it's ALSO a browsable dimension (e.g. "Fact Billing"). */}
                {measuresByGroup.has(name) ? (
                  <Sigma className="h-3.5 w-3.5 shrink-0 text-gold-strong" />
                ) : (
                  <Table2 className="h-3.5 w-3.5 shrink-0 text-muted" />
                )}
                <span className="truncate text-xs font-medium text-ink">{name}</span>
              </button>
              {isExpanded && isLoadingThisTable && (
                <p className="flex items-center gap-1.5 py-1 pl-7 text-[11px] text-muted">
                  <Loader2 className="h-3 w-3 animate-spin" /> Loading columns…
                </p>
              )}
              {isExpanded && !isLoadingThisTable && (
                <ul>
                  <FolderTreeNode
                    node={tree}
                    folderKeyPrefix={name}
                    depth={0}
                    expandedFolders={expandedFolders}
                    toggleFolder={toggleFolder}
                    isPlaced={isPlaced}
                    onFieldCheckToggle={onFieldCheckToggle}
                    onAddHierarchy={onAddHierarchy}
                    onStartDrilldown={onStartDrilldown}
                  />
                  {fields.length === 0 && <li className="py-1 pl-7 text-[11px] text-muted">No fields</li>}
                </ul>
              )}
            </div>
          );
        })}
        {showValuesPseudoField && (
          <div className={visibleTableNames.length > 0 ? "border-t border-slate-100" : ""}>
            <div
              draggable
              onDragStart={(e) => setFieldDragPayload(e, MEASURES_PLACEHOLDER, "available")}
              className="flex cursor-grab items-center gap-1.5 bg-white px-2 py-1.5 font-mono text-[11.5px] text-ink hover:bg-slate-50 active:cursor-grabbing"
            >
              <span className="w-3.5 shrink-0" aria-hidden="true" />
              <span className="truncate">Σ Values</span>
            </div>
          </div>
        )}
        {visibleTableNames.length === 0 && !showValuesPseudoField && (
          <p className="px-2 py-2 text-xs text-muted">No matching fields.</p>
        )}
      </div>
    </div>
  );
}

// A generic drop target for the Values and Filters zones (Rows/Columns use
// HierarchicalCubePathSelect's own built-in drop handling instead — see that component). Visually
// matches the dashed-border "drop here" affordance in HierarchicalCubePathSelect for consistency
// across all four zones.
//
// flex-1: the four zone cards sit in a CSS grid row that stretches every card to match the
// tallest sibling (default grid align-items: stretch), so a short card (e.g. Filters with one
// pill next to a taller Rows/Columns box) ends up with real leftover white space inside it. Without
// flex-1 here, THIS div (the one actually holding onDrop) stayed sized to its own content, leaving
// that leftover space visually part of the card but not draggable-onto — caught live, 2026-09-23.
// Needs its parent card to be a flex column (see FieldPicker's Filters/Values boxes) to have any
// effect; harmless no-op otherwise.
function ZoneDropArea({ zoneKey, onFieldDropped, isFieldAccepted, isEmpty, emptyText = "Drag a field here", emptyIsDropTarget = true, children }) {
  const [dragOver, setDragOver] = useState(false);

  const handleDrop = (e) => {
    e.preventDefault();
    // Stops the drop from bubbling up to PivotBuilder's own page-level "dropped outside every
    // zone = remove the field" fallback (Excel's own behavior: drag a field out of the Field List
    // entirely, anywhere on the page, and it's removed) — this IS a recognized zone, so the drop
    // ends here whether or not it's actually accepted below, same as Excel doing nothing (not
    // removing) when you drop ON a zone but the field doesn't belong there.
    e.stopPropagation();
    setDragOver(false);
    const payload = readFieldDragPayload(e);
    if (payload && payload.source !== zoneKey && (!isFieldAccepted || isFieldAccepted(payload.field))) {
      onFieldDropped(payload.field, payload.source, zoneKey);
    }
  };

  return (
    // min-h applies regardless of content — Excel's own drop zones stay the same tall box whether
    // empty or holding pills, never shrinking to fit just what's inside (flagged live, 2026-10-04:
    // "ليه المربعات الـ 4 لما بحط فيهم حاجة واحدة بيصغروا... مع ان الإكسل مش كدا"). Only the
    // dashed border/padding is conditional on actually being empty.
    // 6.5rem — slightly past the original 6rem (2026-10-12), on request, after restoring from an
    // earlier 2.75rem cut made chasing a page-scrollbar bug that turned out to be caused by
    // AppShell.jsx's own padding instead (now fixed directly there), so none of those cuts were ever
    // actually necessary. Still a real fixed minimum (not shrink-to-content), so the "stays the same
    // size empty or full" Excel parity from the original 2026-10-04 decision holds.
    <div
      className={`min-h-[6.5rem] flex-1 rounded-lg ${dragOver ? "bg-gold/10 ring-2 ring-gold ring-inset" : ""} ${
        isEmpty && emptyIsDropTarget ? "border-2 border-dashed border-slate-200 p-2" : ""
      }`}
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={handleDrop}
    >
      {isEmpty ? <p className="text-xs text-muted">{emptyText}</p> : children}
    </div>
  );
}

// Real Excel's own filter dropdown never mixes these into one flat list — the member checklist is
// always the default view the instant a field is picked, with Label Filters/Value Filters sitting
// as two SEPARATE named flyout groups below it (confirmed live, 2026-10-12, against a real Excel
// screenshot: our old single FILTER_MODE_OPTIONS dropdown combining "Members"/"Exclude" with every
// rule-based mode was flagged as invented UI, "حاسس انك مصمم تجود من عندك"). This now only covers
// the RULE-based modes, grouped to match Excel's own Label/Value Filters submenus — Members/Exclude
// are handled separately below (the checklist is always visible; Exclude is a small toggle beside
// it, not a dropdown entry — see FilterRow's own WHY comment for why that one stays conservative).
const MORE_FILTER_GROUPS = [
  {
    label: "Label Filters",
    options: [
      { value: "LabelContains", label: "Contains…" },
      { value: "LabelBeginsWith", label: "Begins With…" },
      { value: "LabelEndsWith", label: "Ends With…" },
    ],
  },
  {
    label: "Value Filters",
    options: [
      { value: "TopN", label: "Top N…" },
      { value: "BottomN", label: "Bottom N…" },
      { value: "GreaterThan", label: "Greater Than…" },
      { value: "LessThan", label: "Less Than…" },
      { value: "Between", label: "Between…" },
    ],
  },
];
const MORE_FILTER_MODES = new Set(MORE_FILTER_GROUPS.flatMap((g) => g.options.map((o) => o.value)));

// One filter's member list is only fetched once its field is chosen (and only needed for
// Mode "Members"/"Exclude" — the rule-based modes below never show a member list, they describe a
// condition the real cube evaluates itself).
//
// Re-fetches on every search keystroke (debounced) rather than fetching once and filtering
// client-side — the server only ever returns a capped page of members (see
// CubeMetadataService.MaxMembers), so a field past that cap (e.g. "Doctor Name" at 51,512) needs
// the ACTUAL search term sent to the cube each time; searching only within an already-fetched
// unfiltered page can silently miss a real member that just wasn't in that first page at all
// (caught live 2026-09-23).
// Real Excel's own filter dropdown stages EVERY edit (field, mode, text, numbers, the member
// checklist) locally behind one outer OK/Cancel pair — nothing re-queries the cube until OK is
// clicked, and Cancel (or closing any other way) discards the whole draft (confirmed live,
// 2026-10-12, against a real Excel screenshot: our old version called the live onChange on every
// single edit, re-triggering PivotBuilder's auto-run per keystroke/click, with only yesterday's
// MemberCheckboxList-only Apply button as a partial, inconsistent fix). `filter` is only ever read
// ONCE, as this local `draft`'s initial value — Popover.jsx only mounts its children while open
// (see its own `open && (...)` panel), so FilterRow remounts fresh every time this popover opens,
// which is what makes a plain `useState(filter)` initializer correct here with no extra re-sync
// effect needed. `onChange` (the real, live-request-mutating one) is now called exactly once, from
// handleOk — everywhere else in this component reads/writes `draft`/`setDraft` instead.
//
// FilterRow is used in TWO structurally different real-Excel contexts, confirmed live 2026-10-12
// against three separate screenshots — these are two INDEPENDENT axes, not one combined flag:
//
// `hideFieldPicker` (default false): whether the field is already fixed before this popover opens.
// Real Excel NEVER shows a "choose a field" control inside a filter dropdown — the field is always
// already decided by the time you open one (dragged into Report Filter, or it's whichever Row/
// Column header you clicked). This app's OWN "+"-button-adds-an-empty-filter-slot flow used to be
// the one exception, needing an inline picker — fixed 2026-10-12 so "+" now opens a small
// field-picker-only popover FIRST (mirrors the Values zone's own "+" popover), meaning by the time
// FilterRow itself ever mounts, `filter.field` is always already set — so this is passed `true`
// at EVERY call site now, with the old inline Select/drag-handle/remove-X kept only as a defensive
// fallback for the (should-be-impossible) case of a still-fieldless filter reaching here.
//
// `showAdvancedFilters` (default false): which of the two real Excel dropdown SHAPES this is.
// False = Report Filter (page field) shape — ONLY the member checklist + "Select Multiple Items";
// confirmed by a screenshot showing literally nothing between the search box and OK/Cancel besides
// the list and that checkbox (no Label/Value Filters, no Exclude toggle at all). True = a Row/
// Column field's own filter-arrow dropdown shape (FieldHeaderMenu) — confirmed by a screenshot of a
// real "Patient Name" Row field's own dropdown: has "Label Filters"/"Value Filters" but no "Select
// Multiple Items" (a Row/Column field is always unrestricted multi-select, no toggle needed).
// "Exclude these items" has no confirmed real-Excel equivalent in EITHER shape — kept available
// only when showAdvancedFilters (closer to where Label/Value Filters live) rather than removed
// outright, since the backend's Exclude mode is a real, already-shipped feature with no other
// reachable UI entry point yet; flagged as still-unconfirmed, not claimed as a verified match.
export function FilterRow({ filter, dimensionOptions, selectedValueOptions, onChange, onRemove, onClose, hideFieldPicker = false, showAdvancedFilters = false }) {
  const [draft, setDraft] = useState(filter);
  const [memberOptions, setMemberOptions] = useState([]);
  const [loadingMembers, setLoadingMembers] = useState(false);
  // Distinct from "options came back empty" — a failed fetch (e.g. the cube's own well-known
  // "not enough memory" crash, see CubeMetadataService) left this silently indistinguishable from
  // a genuinely empty field ("No items") with no way to tell the two apart, since the catch below
  // only ever re-threw into an unhandled rejection nobody could see (caught live, 2026-10-06: a
  // normally tiny field like BU Short Name showing "No items" during a cube outage, impossible to
  // tell apart from it actually having zero members without reading server logs).
  const [membersFailed, setMembersFailed] = useState(false);
  const [membersRetryToken, setMembersRetryToken] = useState(0);
  const [memberSearch, setMemberSearch] = useState("");
  const mode = draft.mode ?? "Members";
  const isMoreFilterMode = MORE_FILTER_MODES.has(mode);
  // Real Excel's Row/Column field dropdown (2026-10-14): an UNFILTERED field opens with every item
  // checked, and OK with everything checked means "no filter". Here that's an empty Members list,
  // shown as all-checked. noneChecked distinguishes "the user unchecked (Select All)" (nothing
  // checked — OK disabled, like Excel) from that unfiltered state, since both are an empty list.
  const [noneChecked, setNoneChecked] = useState(false);
  const emptyMeansAll = showAdvancedFilters && mode === "Members";
  const showsAllChecked = emptyMeansAll && draft.includedMembers.length === 0 && !noneChecked;

  // A search typed for the PREVIOUS field shouldn't silently carry over and narrow the new
  // field's member list before the user has typed anything for it.
  useEffect(() => {
    setMemberSearch("");
  }, [draft.field]);

  useEffect(() => {
    if (!draft.field || (mode !== "Members" && mode !== "Exclude")) {
      setMemberOptions([]);
      return;
    }
    let cancelled = false;
    const controller = new AbortController();
    setLoadingMembers(true);
    setMembersFailed(false);
    // Debounced: waits for a pause in typing instead of firing one cube query per keystroke,
    // same reasoning as CardinalityGuard exists at all — the shared cube server has crashed
    // before under query load.
    const timeoutId = setTimeout(() => {
      getMembers(draft.field, memberSearch || undefined, controller.signal)
        .then((members) => {
          if (!cancelled) setMemberOptions(members);
        })
        .catch((err) => {
          // AbortError is expected every time a newer keystroke supersedes this one — not a real
          // failure, nothing to show for it.
          if (!cancelled && err.name !== "AbortError") setMembersFailed(true);
        })
        .finally(() => {
          if (!cancelled) setLoadingMembers(false);
        });
    }, 400);
    return () => {
      cancelled = true;
      clearTimeout(timeoutId);
      // Aborts an already-in-flight request (the timer already fired), not just an upcoming one —
      // the backend's own CubeQueryService/CubeMetadataService now actually stop hitting the cube
      // when their request is aborted (see CubeMetadataService.GetMembersAsync's own WHY comment),
      // instead of finishing a search nobody's waiting on anymore (requested live, 2026-10-05).
      controller.abort();
    };
  }, [draft.field, mode, memberSearch, membersRetryToken]);

  const selectedField = dimensionOptions.find((opt) => opt.value === draft.field) ?? null;

  const handleModeChange = (nextMode) => {
    setDraft({ field: draft.field, mode: nextMode, includedMembers: [], labelText: "", n: 10, byMeasureField: null, value: null, valueTo: null });
  };

  const handleOk = () => {
    if (emptyMeansAll && draft.includedMembers.length === 0) {
      // Everything checked = no filter: drop an existing filter rather than save an empty one.
      const hadFilter = (filter.mode ?? "Members") !== "Members" || (filter.includedMembers?.length ?? 0) > 0;
      if (hadFilter) onRemove?.();
      onClose?.();
      return;
    }
    onChange(draft);
    onClose?.();
  };
  const handleCancel = () => onClose?.();

  return (
    <div className="mb-4 rounded-xl border border-slate-200 bg-slate-50/80 p-4">
      {/* Defensive fallback only — every real call site now passes hideFieldPicker=true, since a
          filter's field is always already fixed before FilterRow ever mounts (see this component's
          own WHY comment above). Kept so a still-fieldless filter (should be unreachable) doesn't
          render as a dead end with no way to pick a field or remove itself. */}
      {!hideFieldPicker && (
        <div className="mb-3 flex items-start gap-3">
          {draft.field && (
            <span
              draggable
              onDragStart={(e) => setFieldDragPayload(e, draft.field, "filters")}
              className="mt-2 cursor-grab text-muted hover:text-ink active:cursor-grabbing"
              title="Drag to move this field to another zone"
            >
              <GripVertical className="h-5 w-5" />
            </span>
          )}
          <div className="flex-1">
            <Select
              className="react-select-container"
              classNamePrefix="react-select"
              placeholder="Choose a field to filter…"
              options={dimensionOptions}
              value={selectedField}
              onChange={(opt) => setDraft({ field: opt?.value ?? "", mode: "Members", includedMembers: [] })}
              isClearable
              menuPortalTarget={typeof document !== "undefined" ? document.body : null}
              menuPosition="fixed"
              styles={{ menuPortal: (base) => ({ ...base, zIndex: 10000 }) }}
            />
          </div>
        </div>
      )}
      {draft.field && (
        <div className="flex flex-col gap-3">
          {/* The member checklist (today's "Members" mode) is the ALWAYS-visible default the
              instant a field is picked — no mode dropdown needed to reach it, matching real
              Excel's own filter dropdown. "More Filters" only renders when showAdvancedFilters (the
              Row/Column field context) — confirmed absent from the Report Filter context's own
              dropdown entirely (see this component's own WHY comment above).
              The "Exclude these items" checkbox that used to sit here is gone (2026-10-13) — a full
              screenshot of a real Row field's own filter dropdown (Sort A-Z/Z-A, More Sort Options,
              Clear Filter, Label Filters, Value Filters, search, checklist, OK/Cancel) confirmed it
              has NO such control anywhere; this app's own copy of it was never actually verified
              against Excel and turned out to be invented. The backend's Exclude filter mode still
              exists (ValidateDuplicateValues etc. don't reference it, nothing breaks), it's just no
              longer reachable from this menu — no real Excel UI path means no UI path here either. */}
          {showAdvancedFilters && (
            <select
              className={`${nativeSelectClass} w-full`}
              value={isMoreFilterMode ? mode : ""}
              onChange={(e) => (e.target.value ? handleModeChange(e.target.value) : handleModeChange("Members"))}
            >
              <option value="">{isMoreFilterMode ? "Back to member list" : "More Filters…"}</option>
              {MORE_FILTER_GROUPS.map((group) => (
                <optgroup key={group.label} label={group.label}>
                  {group.options.map((opt) => (
                    <option key={opt.value} value={opt.value}>
                      {opt.label}
                    </option>
                  ))}
                </optgroup>
              ))}
            </select>
          )}

          {(mode === "Members" || mode === "Exclude") && membersFailed && (
            <div className="flex items-center justify-between gap-2 rounded-xl border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">
              <span>Couldn't load members — the cube may be busy.</span>
              <button
                type="button"
                onClick={() => setMembersRetryToken((t) => t + 1)}
                className="shrink-0 rounded-lg border border-rose-300 px-2 py-1 text-xs font-medium hover:bg-rose-100"
              >
                Retry
              </button>
            </div>
          )}
          {(mode === "Members" || mode === "Exclude") && !membersFailed && (
            <MemberCheckboxList
              options={memberOptions}
              // In Exclude mode ("Include new items in manual filter") draft.includedMembers holds
              // the EXCLUDED members, but the checklist must still show what's INCLUDED, exactly
              // like Members mode — so the checked set is inverted here on the way in and out.
              // Caught live 2026-10-14: the list showed excluded members as checked, so unchecking
              // one excluded the OTHER member instead. On the way out, exclusions hidden by the
              // current (server-filtered) search are kept, so searching can't silently drop them.
              value={
                mode === "Exclude" ? memberOptions.filter((m) => !draft.includedMembers.includes(m)) : showsAllChecked ? memberOptions : draft.includedMembers
              }
              onChange={(checked) => {
                if (emptyMeansAll) {
                  // Every member checked again (no search narrowing the list) = back to unfiltered.
                  const allChecked = !memberSearch && memberOptions.length > 0 && memberOptions.every((m) => checked.includes(m));
                  setNoneChecked(checked.length === 0);
                  setDraft((prev) => ({ ...prev, includedMembers: allChecked ? [] : checked }));
                  return;
                }
                setDraft((prev) => {
                  if (mode !== "Exclude") return { ...prev, includedMembers: checked };
                  const visible = new Set(memberOptions);
                  const hiddenExclusions = prev.includedMembers.filter((m) => !visible.has(m));
                  return { ...prev, includedMembers: [...hiddenExclusions, ...memberOptions.filter((m) => !checked.includes(m))] };
                });
              }}
              placeholder={loadingMembers ? "Loading members…" : "Type to search members…"}
              isLoading={loadingMembers}
              showSelectMultiple={!showAdvancedFilters}
              onInputChange={(text, actionMeta) => {
                // Only a real keystroke should update the search term — react-select also fires
                // this on "menu-close"/"input-blur"/"set-value" with a value we don't want (e.g.
                // clearing it to "" right when the user picks an option, which would immediately
                // re-fetch the UNFILTERED list under the still-open menu).
                if (actionMeta.action === "input-change") {
                  setMemberSearch(text);
                }
              }}
            />
          )}

          {/* Real Excel's own "Include new items in manual filter" lives in the Field Settings
              dialog (PLAN 6), not here — placed here instead (PLAN 7, 2026-10-13) since this is
              where the backend's existing Members/Exclude modes (and the full member list needed to
              flip between them) already live; wiring it through the separate Field Settings popover
              would need threading request.filters and a live member fetch into a component that has
              neither today. Unchecked (default) = Members mode: the checked members ARE the
              selection, anything not checked (including a member added to the cube LATER) stays
              excluded. Checked = Exclude mode: the SAME visually-checked set is preserved by
              inverting storage to its complement, so an unseen/future member defaults to INCLUDED
              instead (MdxPivotQueryBuilder's Except(...) already handles an empty exclusion list as
              "everyone" — see ValidateFilterShape's own WHY comment). Needs the FULL member list to
              compute a correct complement, which `memberOptions` only holds while memberSearch is
              empty (it's server-filtered otherwise) — disabled while searching rather than risk
              silently wrong membership. */}
          {(mode === "Members" || mode === "Exclude") && !membersFailed && (
            <label
              className={`flex items-center gap-2 text-xs text-ink ${memberSearch ? "cursor-not-allowed text-muted/60" : "cursor-pointer"}`}
              title={memberSearch ? "Clear the search box first" : undefined}
            >
              <input
                type="checkbox"
                className="h-3.5 w-3.5 accent-gold"
                checked={mode === "Exclude"}
                disabled={Boolean(memberSearch)}
                onChange={(e) => {
                  // An empty Members selection means "no filter yet" (everyone), so its Exclude
                  // equivalent is "exclude nobody" — not the complement, which would exclude
                  // everyone and return an empty pivot.
                  const complement =
                    e.target.checked && mode === "Members" && draft.includedMembers.length === 0
                      ? []
                      : memberOptions.filter((m) => !draft.includedMembers.includes(m));
                  setDraft((prev) => ({ ...prev, mode: e.target.checked ? "Exclude" : "Members", includedMembers: complement }));
                }}
              />
              Include new items in manual filter
            </label>
          )}

          {(mode === "LabelContains" || mode === "LabelBeginsWith" || mode === "LabelEndsWith") && (
            <input
              type="text"
              className="input-field"
              placeholder="Text to match…"
              value={draft.labelText ?? ""}
              onChange={(e) => setDraft((prev) => ({ ...prev, labelText: e.target.value }))}
            />
          )}

          {(mode === "TopN" || mode === "BottomN") && (
            <div className="flex flex-wrap items-center gap-3">
              <input
                type="number"
                min={1}
                className="input-field w-24"
                value={draft.n ?? 10}
                onChange={(e) => setDraft((prev) => ({ ...prev, n: Number(e.target.value) }))}
              />
              <select
                className={nativeSelectClass}
                value={draft.byMeasureField ?? ""}
                onChange={(e) => setDraft((prev) => ({ ...prev, byMeasureField: e.target.value || null }))}
              >
                <option value="">Choose a measure…</option>
                {selectedValueOptions.map((m) => (
                  <option key={m.value} value={m.value}>
                    {m.label}
                  </option>
                ))}
              </select>
            </div>
          )}

          {(mode === "GreaterThan" || mode === "LessThan" || mode === "Between") && (
            <div className="flex flex-wrap items-center gap-3">
              <input
                type="number"
                className="input-field w-28"
                placeholder={mode === "Between" ? "From…" : "Value…"}
                value={draft.value ?? ""}
                onChange={(e) => setDraft((prev) => ({ ...prev, value: e.target.value === "" ? null : Number(e.target.value) }))}
              />
              {mode === "Between" && (
                <input
                  type="number"
                  className="input-field w-28"
                  placeholder="To…"
                  value={draft.valueTo ?? ""}
                  onChange={(e) => setDraft((prev) => ({ ...prev, valueTo: e.target.value === "" ? null : Number(e.target.value) }))}
                />
              )}
              <select
                className={nativeSelectClass}
                value={draft.byMeasureField ?? ""}
                onChange={(e) => setDraft((prev) => ({ ...prev, byMeasureField: e.target.value || null }))}
              >
                <option value="">Choose a measure…</option>
                {selectedValueOptions.map((m) => (
                  <option key={m.value} value={m.value}>
                    {m.label}
                  </option>
                ))}
              </select>
            </div>
          )}
        </div>
      )}
      {/* Outer OK/Cancel — the one place this whole draft actually reaches the live request (via
          handleOk's onChange call). Cancel discards everything typed/checked above without
          touching it, matching real Excel's own dropdown exactly. "Remove filter" moved here (from
          the now-hidden field-picker row's own X) once hideFieldPicker became the normal case —
          still the one place a filter can be deleted outright, not just edited. */}
      <div className="mt-4 flex items-center justify-between gap-2 border-t border-slate-200 pt-3">
        {hideFieldPicker && onRemove ? (
          <button
            type="button"
            onClick={() => {
              onRemove();
              onClose?.();
            }}
            className="text-xs font-medium text-muted hover:text-ink"
          >
            Remove filter
          </button>
        ) : (
          <span />
        )}
        <div className="flex gap-2">
          <button type="button" onClick={handleCancel} className="btn-secondary px-3 py-1.5 text-xs">
            Cancel
          </button>
          {/* Disabled with nothing checked, like Excel's own filter dropdown. */}
          <button
            type="button"
            onClick={handleOk}
            disabled={emptyMeansAll && noneChecked}
            className="btn-primary px-3 py-1.5 text-xs disabled:cursor-not-allowed disabled:opacity-50"
          >
            OK
          </button>
        </div>
      </div>
    </div>
  );
}

// The compact Filters drop-zone box shows just this pill (field name only) — the full config
// (mode, members/text/N, ...) lives in a Popover behind it, reusing FilterRow unchanged. The pill
// itself is draggable so moving a filter to another zone doesn't require opening the popover
// first, same pattern as HierarchicalCubePathSelect's own pills.
function FilterPill({ filter, dimensionOptions, selectedValueOptions, onChange, onRemove, onDropReorder, onMoveTo }) {
  const label = filter.field ? formatDimensionName(filter.field) : "New filter";
  return (
    // onDragOver/onDrop here, not on the <span> below: dropping ON this pill reorders/moves a
    // field onto it specifically (see handleFilterRowDrop) — same wrapper pattern as ValuePill.
    // w-full: the parent's flex-wrap row stacks one full-width pill per line this way, matching
    // Excel's own field list (each field its own row, not several short chips side by side).
    <div className="w-full min-w-0" onDragOver={(e) => e.preventDefault()} onDrop={onDropReorder}>
      {/* align="right": the field pane now sits docked at the page's right edge (see
          PivotBuilder.jsx) — default left-aligned/expand-right would grow off that edge, same fix
          already needed once for ValuePill/Rows before this layout change (2026-09-27). */}
      {/* Excel's own field-button look: a full-width plain bordered button with a dropdown
          chevron, not a colored chip (flagged live, 2026-10-04, against a screenshot of real
          Excel's field list). No right-click menu — a real Excel pill only responds to a LEFT
          click on its own arrow (confirmed live, 2026-10-05); "Move to Rows/Columns" sit as quick
          links at the top of the same popover FilterRow already opens, instead of a second,
          separately-positioned UI. FilterRow has its own "Remove filter" link inside the popover
          itself, so no separate X button is needed here either. */}
      <Popover
        align="right"
        fullWidth
        trigger={(toggle) => (
          <button
            type="button"
            draggable={Boolean(filter.field)}
            onDragStart={(e) => filter.field && setFieldDragPayload(e, filter.field, "filters")}
            onClick={toggle}
            className={`flex w-full items-center justify-between gap-1 rounded border border-slate-300 bg-white px-2 py-1.5 text-left text-xs font-medium text-ink shadow-sm hover:bg-slate-50 ${
              filter.field ? "cursor-grab active:cursor-grabbing" : ""
            }`}
          >
            <span className="min-w-0 truncate">{label}</span>
            <ChevronDown className="h-3.5 w-3.5 shrink-0 text-muted" />
          </button>
        )}
      >
        {(close) => (
          <div className="space-y-2">
            {/* An explicit close button, not just click-outside/Escape — requested live, 2026-10-06
                ("دي ليه مافيش زرار بيقفلها") after the original design relied entirely on implicit
                dismissal (see this component's own earlier WHY comment above, now outdated on that
                point). Real Excel's own filter dialogs always have a visible close affordance too. */}
            <div className="flex items-start justify-between gap-2">
              {filter.field ? (
                <div className="flex flex-wrap gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      onMoveTo("rows");
                      close();
                    }}
                    className="btn-secondary px-2 py-1 text-xs"
                  >
                    Move to Rows
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      onMoveTo("columns");
                      close();
                    }}
                    className="btn-secondary px-2 py-1 text-xs"
                  >
                    Move to Columns
                  </button>
                </div>
              ) : (
                <div />
              )}
              <button type="button" onClick={close} className="shrink-0 rounded-full p-1 text-muted hover:bg-slate-100 hover:text-ink" aria-label="Close">
                <X className="h-4 w-4" />
              </button>
            </div>
            {filter.field && <div className="border-b border-slate-100" />}
            <FilterRow
              filter={filter}
              dimensionOptions={dimensionOptions}
              selectedValueOptions={selectedValueOptions}
              onChange={onChange}
              onRemove={onRemove}
              onClose={close}
              hideFieldPicker={Boolean(filter.field)}
            />
          </div>
        )}
      </Popover>
    </div>
  );
}

// Excel's "Group Field" (date) and "Group Selection" (numeric bins), one control per Rows OR
// Columns field (2026-09-24: extended from Rows-only) — backend enforces one grouping per field
// (see MdxPivotQueryBuilder.ValidateGroupings). Mode is derived from whichever of
// dateGroupings/numericGroupings currently mentions the field, rather than stored as its own
// separate piece of state, so the two arrays stay the single source of truth the backend actually
// reads. Rendered once per axis (see the two call sites below) — dateGroupings/numericGroupings
// are shared arrays across both, since a grouping object only names its field, not an axis.
// Real Excel's own "Group Field" is simply unavailable for a plain text field (e.g. a Business
// Unit code) — only a date or numeric field can ever be grouped — but this control used to render
// a "No grouping" dropdown for EVERY Rows/Columns field regardless of its actual type, directly
// contradicting its own "(date/numeric fields only)" caption (caught live, 2026-10-06: "هي دي انت
// عاملها ليه انا مش لاقيها في الـ excel هل هي موجودة؟" against a BU Short Name field, which is
// neither). `dimensions` (CubeField.IsNumeric/IsDate, from the cube's own LEVEL_DBTYPE — see the
// backend's own WHY comment) is what finally lets this filter down to fields that can actually use
// it, and restrict EACH field's own dropdown to the options that make sense for ITS type (a numeric
// field never gets Month/Quarter/Year, a date field never gets Numeric bins — real Excel draws that
// same line, "Group Field" vs "Group Selection" are never offered together for the same field).
function AxisGroupingEditor({ label, fields, dimensions, dateGroupings, numericGroupings, onChange }) {
  const groupableFields = fields.filter((field) => {
    const dim = dimensions.find((d) => d.field === field);
    return dim?.isNumeric || dim?.isDate;
  });
  if (groupableFields.length === 0) {
    return null;
  }
  const fieldIsNumeric = (field) => dimensions.find((d) => d.field === field)?.isNumeric ?? false;

  const modeFor = (field) => {
    const dateGrouping = dateGroupings.find((g) => g.field === field);
    if (dateGrouping) return dateGrouping.unit;
    const numericGrouping = numericGroupings.find((g) => g.field === field);
    if (numericGrouping) return "Numeric";
    return "None";
  };

  const setMode = (field, mode) => {
    const nextDate = dateGroupings.filter((g) => g.field !== field);
    const nextNumeric = numericGroupings.filter((g) => g.field !== field);
    if (mode === "Month" || mode === "Quarter" || mode === "Year") {
      nextDate.push({ field, unit: mode });
    } else if (mode === "Numeric") {
      nextNumeric.push({ field, binSize: 10 });
    }
    onChange(nextDate, nextNumeric);
  };

  const setBinSize = (field, binSize) => {
    onChange(
      dateGroupings,
      numericGroupings.map((g) => (g.field === field ? { ...g, binSize } : g)),
    );
  };

  return (
    <div className="mt-4 space-y-2">
      <p className="text-xs font-medium text-muted">{label} (date/numeric fields only):</p>
      {groupableFields.map((field) => {
        const mode = modeFor(field);
        const isNumeric = fieldIsNumeric(field);
        return (
          <div key={field} className="flex flex-wrap items-center gap-2">
            <span className="min-w-[10rem] text-sm text-ink">{formatDimensionName(field)}</span>
            <select className={nativeSelectClass} value={mode} onChange={(e) => setMode(field, e.target.value)}>
              <option value="None">No grouping</option>
              {isNumeric ? (
                <option value="Numeric">Numeric bins</option>
              ) : (
                <>
                  <option value="Month">Group by Month</option>
                  <option value="Quarter">Group by Quarter</option>
                  <option value="Year">Group by Year</option>
                </>
              )}
            </select>
            {mode === "Numeric" && (
              <input
                type="number"
                min={0.01}
                step="any"
                className="input-field w-24 py-2"
                value={numericGroupings.find((g) => g.field === field)?.binSize ?? 10}
                onChange={(e) => setBinSize(field, Number(e.target.value))}
                title="Bin size"
              />
            )}
          </div>
        );
      })}
    </div>
  );
}

// Custom order is Excel's manual drag-to-reorder — an explicit list of members in the exact
// display order, rather than an automatic rule. react-select's multi-value order already follows
// click order (and reordering is just remove-then-reclick), so reusing it here for BUILDING the
// list doubles as the reordering UI without a separate drag-and-drop implementation.
function CustomOrderEditor({ field, customOrder, onChange }) {
  const [memberOptions, setMemberOptions] = useState([]);
  const [loadingMembers, setLoadingMembers] = useState(false);
  // See FilterRow's own membersFailed WHY comment — same silent-failure gap, same fix, every place
  // this file calls getMembers.
  const [membersFailed, setMembersFailed] = useState(false);
  const [membersRetryToken, setMembersRetryToken] = useState(0);

  useEffect(() => {
    if (!field) {
      setMemberOptions([]);
      return;
    }
    let cancelled = false;
    const controller = new AbortController();
    setLoadingMembers(true);
    setMembersFailed(false);
    getMembers(field, undefined, controller.signal)
      .then((members) => {
        if (!cancelled) setMemberOptions(members);
      })
      .catch((err) => {
        if (!cancelled && err.name !== "AbortError") setMembersFailed(true);
      })
      .finally(() => {
        if (!cancelled) setLoadingMembers(false);
      });
    return () => {
      cancelled = true;
      controller.abort();
    };
  }, [field, membersRetryToken]);

  if (!field) {
    return <p className="text-sm text-muted">Add exactly one field to the target axis first to set a custom order.</p>;
  }

  if (membersFailed) {
    return (
      <div className="flex items-center justify-between gap-2 rounded-xl border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">
        <span>Couldn't load members — the cube may be busy.</span>
        <button
          type="button"
          onClick={() => setMembersRetryToken((t) => t + 1)}
          className="shrink-0 rounded-lg border border-rose-300 px-2 py-1 text-xs font-medium hover:bg-rose-100"
        >
          Retry
        </button>
      </div>
    );
  }

  return (
    <MultiSelectField
      options={memberOptions}
      value={customOrder ?? []}
      onChange={onChange}
      placeholder={loadingMembers ? "Loading members…" : "Click members in the order you want them displayed…"}
      isLoading={loadingMembers}
    />
  );
}

// One Rows/Columns FIELD's own independent sort rule (Excel's real per-field "Sort A to
// Z"/"Sort Largest to Smallest") — mutually exclusive with the single axis-wide Sort above (see
// LevelSort's own WHY comment on the backend model for why this is what finally makes
// measure-sort compatible with Subtotals). Only shown once an axis has 2+ fields — with exactly
// one field, the single Sort dropdown above already covers both label and by-value sorting.
function LevelSortEditor({ label, fields, levelSorts, values, onChange }) {
  if (fields.length < 2) {
    return null;
  }

  const modeFor = (level) => {
    const rule = levelSorts.find((s) => s.level === level);
    return rule ? (rule.byMeasureField ?? "label") : "none";
  };
  const directionFor = (level) => levelSorts.find((s) => s.level === level)?.direction ?? "Ascending";

  const setRule = (level, mode, direction) => {
    const withoutLevel = levelSorts.filter((s) => s.level !== level);
    if (mode === "none") {
      onChange(withoutLevel);
      return;
    }
    onChange([...withoutLevel, { level, direction, byMeasureField: mode === "label" ? null : mode }]);
  };

  return (
    <div className="mt-3 space-y-2 border-t border-slate-200 pt-3">
      <p className="text-xs font-medium text-muted">{label}:</p>
      {fields.map((field, level) => {
        // Σ Values isn't a real field — no label/value to sort this level by (the backend rejects
        // a LevelSort targeting its position outright, see ValidateAxisLevelSorts) — skipped here
        // rather than filtered out of `fields` itself, since `level` must stay aligned with this
        // field's REAL position in Rows/Columns for every OTHER row's own level value to stay
        // correct.
        if (field === MEASURES_PLACEHOLDER) {
          return null;
        }
        const mode = modeFor(level);
        const direction = directionFor(level);
        return (
          <div key={level} className="flex flex-wrap items-center gap-2">
            <span className="min-w-[9rem] text-sm text-ink">{formatDimensionName(field)}</span>
            <select className={nativeSelectClass} value={mode} onChange={(e) => setRule(level, e.target.value, direction)}>
              <option value="none">No sorting</option>
              <option value="label">By label</option>
              {values.filter((v) => v.field).map((v) => (
                <option key={v.field} value={v.field}>
                  By value of {v.field}
                </option>
              ))}
            </select>
            {mode !== "none" && (
              <select className={nativeSelectClass} value={direction} onChange={(e) => setRule(level, mode, e.target.value)}>
                <option value="Ascending">Ascending</option>
                <option value="Descending">Descending</option>
              </select>
            )}
          </div>
        );
      })}
    </div>
  );
}

// onChange here takes the BUNDLED { sort, rowLevelSorts, columnLevelSorts } — all three interact
// (Sort and *LevelSorts are mutually exclusive, enforced by clearing one when the other is set)
// so they're easiest to manage together rather than as separate onChange callbacks.
export function SortEditor({ sort, rowLevelSorts, columnLevelSorts, values, rows, columns, onChange }) {
  const mode = sort === null ? "none" : sort.customOrder ? "custom" : sort.byMeasureField ? sort.byMeasureField : "label";
  const axis = sort?.axis ?? "Rows";
  // Σ Values isn't a real field — "sort/custom-order by its label" has no meaning, so it's
  // excluded when picking which single field this axis-wide sort actually targets.
  const targetField = (axis === "Columns" ? columns : rows).filter((f) => f !== MEASURES_PLACEHOLDER)[0] ?? null;

  const handleModeChange = (nextMode) => {
    // Setting the single axis-wide sort clears ALL level-by-level rules, on BOTH axes — the backend
    // rejects `sort` whenever EITHER RowLevelSorts or ColumnLevelSorts is non-empty
    // (ValidateLevelSorts), not just the same axis's (PivotBuilder's own default-sort-on-first-field
    // can leave the OTHER axis's LevelSorts populated too, caught live 2026-10-06).
    const clearLevelSorts = { rowLevelSorts: [], columnLevelSorts: [] };
    if (nextMode === "none") {
      onChange({ sort: null });
      return;
    }
    const direction = sort?.direction ?? "Ascending";
    if (nextMode === "custom") {
      onChange({ sort: { direction, byMeasureField: null, axis, customOrder: [] }, ...clearLevelSorts });
      return;
    }
    onChange({
      sort: nextMode === "label" ? { direction, byMeasureField: null, axis, customOrder: null } : { direction, byMeasureField: nextMode, axis, customOrder: null },
      ...clearLevelSorts,
    });
  };

  return (
    <div className="flex flex-col gap-3">
      <div className="flex flex-wrap items-center gap-3">
        <select className={nativeSelectClass} value={mode} onChange={(e) => handleModeChange(e.target.value)}>
          <option value="none">No sorting</option>
          <option value="label">By label (single field only)</option>
          <option value="custom">Custom order (single field only)</option>
          {values.filter((v) => v.field).map((v) => (
            <option key={v.field} value={v.field}>
              By value of {v.field}
            </option>
          ))}
        </select>
        {sort && (
          <>
            <select
              className={nativeSelectClass}
              value={axis}
              onChange={(e) => onChange({ sort: { ...sort, axis: e.target.value } })}
            >
              <option value="Rows">Sort Rows</option>
              <option value="Columns">Sort Columns</option>
            </select>
            {mode !== "custom" && (
              <select
                className={nativeSelectClass}
                value={sort.direction}
                onChange={(e) => onChange({ sort: { ...sort, direction: e.target.value } })}
              >
                <option value="Ascending">Ascending</option>
                <option value="Descending">Descending</option>
              </select>
            )}
          </>
        )}
      </div>
      {sort && mode === "custom" && (
        <CustomOrderEditor
          field={targetField}
          customOrder={sort.customOrder}
          onChange={(customOrder) => onChange({ sort: { ...sort, customOrder } })}
        />
      )}
      {!sort && (
        <>
          <LevelSortEditor
            label="Or sort each Row level independently"
            fields={rows}
            levelSorts={rowLevelSorts}
            values={values}
            onChange={(next) => onChange({ rowLevelSorts: next })}
          />
          <LevelSortEditor
            label="Or sort each Column level independently"
            fields={columns}
            levelSorts={columnLevelSorts}
            values={values}
            onChange={(next) => onChange({ columnLevelSorts: next })}
          />
        </>
      )}
    </div>
  );
}

// A calculated field is a NEW measure derived from two EXISTING measures with a fixed operator —
// not a free-text formula. See the backend's CalculatedField for why (avoids accepting arbitrary
// MDX from the client). Once added, its name becomes selectable in the Values picker below.
export function CalculatedFieldsEditor({ calculatedFields, measureOptions, onChange }) {
  const update = (index, patch) => {
    const copy = [...calculatedFields];
    copy[index] = { ...copy[index], ...patch };
    onChange(copy);
  };

  const remove = (index) => onChange(calculatedFields.filter((_, i) => i !== index));

  const add = () =>
    onChange([...calculatedFields, { name: "", leftField: "", operator: "Divide", rightField: "" }]);

  return (
    <>
      {calculatedFields.map((field, index) => (
        <div key={index} className="mb-3 flex flex-wrap items-center gap-2">
          <input
            type="text"
            placeholder="New field name…"
            value={field.name}
            onChange={(e) => update(index, { name: e.target.value })}
            className="input-field min-w-[10rem] flex-1 py-2"
          />
          <select className={nativeSelectClass} value={field.leftField} onChange={(e) => update(index, { leftField: e.target.value })}>
            <option value="">Choose measure…</option>
            {measureOptions.map((m) => (
              <option key={m.value} value={m.value}>
                {m.label}
              </option>
            ))}
          </select>
          <select className={nativeSelectClass} value={field.operator} onChange={(e) => update(index, { operator: e.target.value })}>
            <option value="Add">+</option>
            <option value="Subtract">−</option>
            <option value="Multiply">×</option>
            <option value="Divide">÷</option>
          </select>
          <select className={nativeSelectClass} value={field.rightField} onChange={(e) => update(index, { rightField: e.target.value })}>
            <option value="">Choose measure…</option>
            {measureOptions.map((m) => (
              <option key={m.value} value={m.value}>
                {m.label}
              </option>
            ))}
          </select>
          <button type="button" onClick={() => remove(index)} className="text-muted hover:text-ink" aria-label="Remove calculated field">
            <X className="h-5 w-5" />
          </button>
        </div>
      ))}
      <button type="button" onClick={add} className="btn-secondary">
        + Add calculated field
      </button>
    </>
  );
}

// Excel's "Calculated Item" — a NEW MEMBER within an EXISTING Rows/Columns field's hierarchy
// (e.g. "East + West" as a new Region member), distinct from CalculatedFieldsEditor above (which
// creates a new MEASURE instead). Same "reject a free-text formula" philosophy as calculated
// fields: a signed sum of the field's own EXISTING members, picked from a member list rather than
// typed — see the backend's CalculatedItem for the full WHY and its v1 scope cuts (can't combine
// with a Filter or grouping on the same field, or with any ShowValuesAs besides Normal).
function CalculatedItemRow({ item, axisFieldOptions, onChange, onRemove }) {
  const [memberOptions, setMemberOptions] = useState([]);
  const [loadingMembers, setLoadingMembers] = useState(false);
  // See FilterRow's own membersFailed WHY comment — same silent-failure gap, same fix. Shared by
  // both the Add(+) and Subtract(−) lists below since they both read from this one fetch.
  const [membersFailed, setMembersFailed] = useState(false);
  const [membersRetryToken, setMembersRetryToken] = useState(0);

  useEffect(() => {
    if (!item.field) {
      setMemberOptions([]);
      return;
    }
    let cancelled = false;
    const controller = new AbortController();
    setLoadingMembers(true);
    setMembersFailed(false);
    getMembers(item.field, undefined, controller.signal)
      .then((members) => {
        if (!cancelled) setMemberOptions(members);
      })
      .catch((err) => {
        if (!cancelled && err.name !== "AbortError") setMembersFailed(true);
      })
      .finally(() => {
        if (!cancelled) setLoadingMembers(false);
      });
    return () => {
      cancelled = true;
      controller.abort();
    };
  }, [item.field, membersRetryToken]);

  return (
    <div className="mb-4 rounded-xl border border-slate-200 bg-slate-50/80 p-4">
      <div className="mb-3 flex flex-wrap items-center gap-2">
        <select
          className={nativeSelectClass}
          value={item.field}
          onChange={(e) => onChange({ field: e.target.value, positiveMembers: [], negativeMembers: [] })}
        >
          <option value="">Choose a Rows/Columns field…</option>
          {axisFieldOptions.map((f) => (
            <option key={f} value={f}>
              {formatDimensionName(f)}
            </option>
          ))}
        </select>
        <input
          type="text"
          placeholder="New item name (e.g. East + West)…"
          value={item.name}
          onChange={(e) => onChange({ name: e.target.value })}
          className="input-field min-w-[10rem] flex-1 py-2"
        />
        <button type="button" onClick={onRemove} className="text-muted hover:text-ink" aria-label="Remove calculated item">
          <X className="h-5 w-5" />
        </button>
      </div>
      {item.field && membersFailed && (
        <div className="flex items-center justify-between gap-2 rounded-xl border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">
          <span>Couldn't load members — the cube may be busy.</span>
          <button
            type="button"
            onClick={() => setMembersRetryToken((t) => t + 1)}
            className="shrink-0 rounded-lg border border-rose-300 px-2 py-1 text-xs font-medium hover:bg-rose-100"
          >
            Retry
          </button>
        </div>
      )}
      {item.field && !membersFailed && (
        <div className="flex flex-col gap-2">
          <div>
            <span className="mb-1 block text-xs font-medium text-muted">Add (+):</span>
            <MemberCheckboxList
              options={memberOptions}
              value={item.positiveMembers}
              onChange={(positiveMembers) => onChange({ positiveMembers })}
              placeholder={loadingMembers ? "Loading members…" : "Select member(s) to add…"}
              isLoading={loadingMembers}
            />
          </div>
          <div>
            <span className="mb-1 block text-xs font-medium text-muted">Subtract (−):</span>
            <MemberCheckboxList
              options={memberOptions}
              value={item.negativeMembers}
              onChange={(negativeMembers) => onChange({ negativeMembers })}
              placeholder={loadingMembers ? "Loading members…" : "Select member(s) to subtract…"}
              isLoading={loadingMembers}
            />
          </div>
        </div>
      )}
    </div>
  );
}

export function CalculatedItemsEditor({ calculatedItems, rows, columns, onChange }) {
  // Σ Values isn't a real field — a Calculated Item (a new synthetic MEMBER of an existing
  // field's hierarchy) has nothing to attach to on it, so it's excluded here even when present in
  // Rows/Columns.
  const axisFieldOptions = [...rows, ...columns].filter((f) => f !== MEASURES_PLACEHOLDER);

  const update = (index, patch) => {
    const copy = [...calculatedItems];
    copy[index] = { ...copy[index], ...patch };
    onChange(copy);
  };

  const remove = (index) => onChange(calculatedItems.filter((_, i) => i !== index));

  const add = () =>
    onChange([...calculatedItems, { field: axisFieldOptions[0] ?? "", name: "", positiveMembers: [], negativeMembers: [] }]);

  if (axisFieldOptions.length === 0) {
    return <p className="text-sm text-muted">Add at least one Row or Column field first.</p>;
  }

  return (
    <>
      {calculatedItems.map((item, index) => (
        <CalculatedItemRow
          key={index}
          item={item}
          axisFieldOptions={axisFieldOptions}
          onChange={(patch) => update(index, patch)}
          onRemove={() => remove(index)}
        />
      ))}
      <button type="button" onClick={add} className="btn-secondary">
        + Add calculated item
      </button>
    </>
  );
}


const NUMBER_FORMAT_CATEGORIES = [
  "General",
  "Number",
  "Currency",
  "Accounting",
  "Percentage",
  "Date",
  "Time",
  "Fraction",
  "Scientific",
  "Text",
  "Special",
  "Custom",
];

// Real Excel's own Format Cells → Currency/Accounting "Symbol" is a DROPDOWN of locale-specific
// currency formats pulled from Windows' own installed regional settings (hundreds of entries,
// several per currency for each language variant — flagged live, 2026-10-06, against a screenshot
// of that real dropdown). Replicating that exact OS-locale list isn't practical or useful here (it's
// Windows' own locale database, not something this app has access to or a reason to reproduce one-
// for-one) — this is a curated list of real currencies relevant to this app's actual users instead,
// plus "Custom…" for anything not listed (free-text, same as the single input this replaces).
const CURRENCY_OPTIONS = [
  { value: "", label: "(None)" },
  { value: "EGP", label: "EGP — Egyptian Pound" },
  { value: "$", label: "$ — US Dollar" },
  { value: "€", label: "€ — Euro" },
  { value: "£", label: "£ — British Pound" },
  { value: "SAR", label: "SAR — Saudi Riyal" },
  { value: "AED", label: "AED — UAE Dirham" },
  { value: "KWD", label: "KWD — Kuwaiti Dinar" },
  { value: "QAR", label: "QAR — Qatari Riyal" },
  { value: "OMR", label: "OMR — Omani Rial" },
  { value: "BHD", label: "BHD — Bahraini Dinar" },
  { value: "JOD", label: "JOD — Jordanian Dinar" },
  { value: "¥", label: "¥ — Japanese Yen / Chinese Yuan" },
  { value: "₹", label: "₹ — Indian Rupee" },
  { value: "₺", label: "₺ — Turkish Lira" },
  { value: "₽", label: "₽ — Russian Ruble" },
  { value: "CHF", label: "CHF — Swiss Franc" },
  { value: "CAD", label: "CAD — Canadian Dollar" },
  { value: "AUD", label: "AUD — Australian Dollar" },
];

// Excel's own 4 "Special" presets (Format Cells → Special) — each is a fixed digit-mask, not a
// distinct numeric transform, so the picker here is just a label/value list for NumberFormat's own
// specialType (see NumberFormatter.cs's FormatAsSpecial for how each mask is actually applied).
const SPECIAL_FORMAT_TYPES = [
  { value: "ZipCode", label: "Zip Code" },
  { value: "ZipCodePlus4", label: "Zip Code + 4" },
  { value: "PhoneNumber", label: "Phone Number" },
  { value: "SocialSecurityNumber", label: "Social Security Number" },
];

// Representative example the Sample box formats live as the user changes settings — Excel's own
// Format Cells dialog previews against whatever value is already in the selected cell; this popup
// has no single cell to point at (it configures the WHOLE field, every row), so a fixed example
// plays that same role instead.
const NUMBER_FORMAT_SAMPLE_VALUE = 1234.5;

// Excel's real "Format Cells" dialog (opened via Value Field Settings' own "Number Format…"
// button) — a Sample preview above a Category list, not a flat inline dropdown (extended
// 2026-10-05, against screenshots of that real dialog). Shares the exact same formatNumber()
// utils/numberFormat.js already uses for the live grid, so the preview here is never out of sync
// with what the grid will actually show.
export function NumberFormatEditor({ format, onUpdateFormat, onBack }) {
  const needsDecimals = !["General", "Date", "Time", "Fraction", "Text", "Special", "Custom"].includes(format.type);
  const needsCurrencySymbol = format.type === "Currency" || format.type === "Accounting";
  const needsSpecialType = format.type === "Special";
  const needsCustomCode = format.type === "Custom";
  // Explicit "Custom…" picked from the Symbol dropdown below overrides the auto-detected state —
  // without this, picking "Custom…" while the symbol is still e.g. "" (None) would immediately
  // snap back to showing "(None)" selected (since "" IS a known preset value), never actually
  // revealing the free-text box to type something new into.
  const [forceCustomSymbol, setForceCustomSymbol] = useState(false);
  const isCustomSymbol = forceCustomSymbol || !CURRENCY_OPTIONS.some((opt) => opt.value === format.currencySymbol);

  return (
    <div className="space-y-3">
      <button type="button" onClick={onBack} className="inline-flex items-center gap-1 text-xs font-medium text-muted hover:text-ink">
        <ChevronLeft className="h-3.5 w-3.5" /> Back
      </button>

      <div>
        <span className="mb-1 block text-xs font-medium text-muted">Sample</span>
        <div className="rounded-md border border-slate-200 bg-slate-50 px-2 py-1.5 text-sm text-ink">
          {formatNumber(NUMBER_FORMAT_SAMPLE_VALUE, format)}
        </div>
      </div>

      <span className="block text-xs font-medium text-muted">Category:</span>
      <div className="flex gap-3">
        <ul className="max-h-48 w-28 shrink-0 overflow-y-auto rounded-md border border-slate-200">
          {NUMBER_FORMAT_CATEGORIES.map((category) => (
            <li key={category}>
              <button
                type="button"
                onClick={() => onUpdateFormat({ type: category })}
                className={`block w-full px-2 py-1.5 text-left text-xs ${
                  format.type === category ? "bg-gold text-white" : "text-ink hover:bg-slate-100"
                }`}
              >
                {category}
              </button>
            </li>
          ))}
        </ul>

        <div className="flex-1 space-y-2">
          {needsDecimals && (
            <div>
              <span className="mb-1 block text-xs font-medium text-muted">Decimal places</span>
              <input
                type="number"
                min={0}
                max={10}
                value={format.decimalPlaces}
                onChange={(e) => onUpdateFormat({ decimalPlaces: Number(e.target.value) })}
                className="input-field w-20 py-1.5 text-sm"
              />
            </div>
          )}
          {needsCurrencySymbol && (
            <div>
              <span className="mb-1 block text-xs font-medium text-muted">Symbol</span>
              {/* w-full explicitly — unlike every other control in this panel (the Decimal Places
                  input, the Category list), a plain <select> with no width class sizes itself to its
                  OWN selected option's text, not its container. "(None)"/short currency codes never
                  exposed this; a longer label like "EGP — Egyptian Pound" pushed the box wider than
                  this narrow flex-1 column and out past the panel's own right edge (caught live,
                  2026-10-06, "دي بايظة"). */}
              <select
                className={`${nativeSelectClass} w-full`}
                value={isCustomSymbol ? "__custom__" : format.currencySymbol}
                onChange={(e) => {
                  if (e.target.value === "__custom__") {
                    setForceCustomSymbol(true);
                  } else {
                    setForceCustomSymbol(false);
                    onUpdateFormat({ currencySymbol: e.target.value });
                  }
                }}
              >
                {CURRENCY_OPTIONS.map((opt) => (
                  <option key={opt.value} value={opt.value}>
                    {opt.label}
                  </option>
                ))}
                <option value="__custom__">Custom…</option>
              </select>
              {isCustomSymbol && (
                <input
                  type="text"
                  value={format.currencySymbol}
                  onChange={(e) => onUpdateFormat({ currencySymbol: e.target.value })}
                  placeholder="Type a symbol…"
                  className="input-field mt-1.5 w-full py-1.5 text-sm"
                />
              )}
            </div>
          )}
          {needsSpecialType && (
            <div>
              <span className="mb-1 block text-xs font-medium text-muted">Type</span>
              {/* Same narrow flex-1 column, same no-width-class bug as the Symbol select above —
                  SPECIAL_FORMAT_TYPES labels (e.g. "Social Security Number") are long enough to
                  trigger it too. */}
              <select
                className={`${nativeSelectClass} w-full`}
                value={format.specialType ?? "ZipCode"}
                onChange={(e) => onUpdateFormat({ specialType: e.target.value })}
              >
                {SPECIAL_FORMAT_TYPES.map((opt) => (
                  <option key={opt.value} value={opt.value}>
                    {opt.label}
                  </option>
                ))}
              </select>
            </div>
          )}
          {needsCustomCode && (
            <div>
              <span className="mb-1 block text-xs font-medium text-muted">Type the number format code</span>
              <input
                type="text"
                value={format.customFormatCode ?? ""}
                onChange={(e) => onUpdateFormat({ customFormatCode: e.target.value })}
                placeholder="#,##0.00"
                className="input-field w-full py-1.5 text-sm"
              />
            </div>
          )}
          {!needsDecimals && !needsCurrencySymbol && !needsSpecialType && !needsCustomCode && (
            <p className="text-xs text-muted">No further options for this category.</p>
          )}
        </div>
      </div>
    </div>
  );
}

// Excel's own Color Scale/Data Bar dialogs let you pick your own colors instead of being stuck
// with the one stock preset — exported so PivotGrid.jsx's value-cell context menu can open the
// exact same editor (reached there via its own "Conditional Formatting…" item) instead of
// duplicating this JSX a second time, same reuse pattern as ValueFieldSettingsPanel/
// NumberFormatEditor above. `onUpdate` here always means "merge this patch into the conditionalFormat
// object" — the caller's own onUpdate wraps that into however IT stores conditionalFormat.
export function ConditionalFormatEditor({ format, onUpdate, onBack }) {
  const type = format?.type ?? "None";
  const showScaleColors = type === "ColorScale" || type === "IconSet";

  return (
    <div className="space-y-3">
      <button type="button" onClick={onBack} className="mb-1 inline-flex items-center gap-1 text-xs font-medium text-muted hover:text-ink">
        <ChevronLeft className="h-3.5 w-3.5" /> Back
      </button>
      <div>
        <span className="mb-1 block text-xs font-medium text-muted">Format style</span>
        <select className={`${nativeSelectClass} w-full`} value={type} onChange={(e) => onUpdate({ type: e.target.value })}>
          <option value="None">None</option>
          <option value="ColorScale">Color scale</option>
          <option value="DataBar">Data bar</option>
          <option value="IconSet">Icon set (traffic lights)</option>
        </select>
      </div>
      {showScaleColors && (
        <div className="grid grid-cols-3 gap-2">
          <div>
            <span className="mb-1 block text-xs font-medium text-muted">Min</span>
            <input
              type="color"
              value={format?.minColor || DEFAULT_MIN_COLOR}
              onChange={(e) => onUpdate({ minColor: e.target.value })}
              className="h-8 w-full cursor-pointer rounded border border-slate-200"
            />
          </div>
          <div>
            <span className="mb-1 block text-xs font-medium text-muted">Mid</span>
            <input
              type="color"
              value={format?.midColor || DEFAULT_MID_COLOR}
              onChange={(e) => onUpdate({ midColor: e.target.value })}
              className="h-8 w-full cursor-pointer rounded border border-slate-200"
            />
          </div>
          <div>
            <span className="mb-1 block text-xs font-medium text-muted">Max</span>
            <input
              type="color"
              value={format?.maxColor || DEFAULT_MAX_COLOR}
              onChange={(e) => onUpdate({ maxColor: e.target.value })}
              className="h-8 w-full cursor-pointer rounded border border-slate-200"
            />
          </div>
        </div>
      )}
      {type === "DataBar" && (
        <div>
          <span className="mb-1 block text-xs font-medium text-muted">Bar color</span>
          <input
            type="color"
            value={format?.barColor || DEFAULT_BAR_COLOR}
            onChange={(e) => onUpdate({ barColor: e.target.value })}
            className="h-8 w-20 cursor-pointer rounded border border-slate-200"
          />
        </div>
      )}
    </div>
  );
}

// The compact Values drop-zone box shows just this pill (measure name only) — which measure it is
// gets set once, from the multi-select above the pill list (see FieldPicker's own render); the
// pill itself only carries per-measure CONFIG (Show Values As, Format, Conditional format) behind
// a Popover, plus drag-to-reorder within the Values zone (onDropReorder), same pattern as
// FilterPill.
function ValuePill({
  v,
  onUpdate,
  onUpdateFormat,
  onRemove,
  onDropReorder,
  rows,
  onMoveTo,
  onMoveUp,
  onMoveDown,
  onMoveToBeginning,
  onMoveToEnd,
  isFirst,
  isLast,
}) {
  const isDimensionField = v.field.startsWith("[");
  const dimensionAggregation = v.dimensionAggregation ?? "Count";
  // Excel's own distinction (see its real Value Field Settings dialog): "Source Name" is always
  // just the plain underlying field, "Age" — the aggregation-qualified "Sum of Age" is the
  // SEPARATE, editable "Custom Name", which defaults to that computed label until overridden.
  const sourceName = isDimensionField ? formatDimensionName(v.field) : v.field;
  // A dimension field (not a real cube measure) dropped into Values defaults to "<Aggregation> of
  // <field>" — Sum for a numeric field, Count otherwise (see moveField's own WHY comment) — and
  // can be changed via "Summarize Values By" below. Mirrors the backend's own naming exactly (see
  // MdxPivotQueryBuilder.BuildMeasureExpression).
  const computedLabel = isDimensionField ? `${dimensionAggregation} of ${sourceName}` : sourceName;
  const label = v.customLabel || computedLabel;
  // Real Excel's own pill only responds to a LEFT click on its dropdown arrow for this menu —
  // there's no separate right-click path at all (confirmed live, 2026-10-05: "هل في الإكسل بندوس
  // right click على اي حاجة؟"). The previous version's right-click PillContextMenu duplicated the
  // exact same items as a second, independently-positioned UI — removed entirely in favor of this
  // one Popover, which just swaps its own content between "menu" and "settings" (same two-state
  // pattern as PivotGrid's FieldHeaderMenu "More Sort Options…").
  const [panel, setPanel] = useState("menu");

  return (
    // w-full: the parent's flex-wrap row stacks one full-width pill per line this way, matching
    // Excel's own field list (each field its own row, not several short chips side by side).
    <div className="w-full min-w-0" onDragOver={(e) => e.preventDefault()} onDrop={onDropReorder}>
      {/* align="right": the Values card sits flush against the right edge of the "Fields"
          strip (it's a w-fit card pushed there by Filters/Columns/Rows' flex-1 — see FieldPicker's
          own WHY comment on that grid), so Popover's default left-aligned/expand-right panel had
          nowhere to expand into and was clipped by the browser's edge (caught live, 2026-09-27).
          Opening it expanding leftward, into the room Filters/Columns/Rows have, keeps it on
          screen. */}
      {/* Excel's own field-button look: a full-width plain bordered button with a dropdown
          chevron, not a colored chip (flagged live, 2026-10-04, against a screenshot of real
          Excel's field list). No separate X here — removal is drag-out-to-remove (already
          supported) or the right-click menu's "Remove Field", matching Excel's own dialog, which
          has no remove control inside it either. */}
      <Popover
        align="right"
        fullWidth
        panelClassName={panel === "menu" ? "!min-w-[13rem] !max-w-[16rem] !p-1" : ""}
        trigger={(toggle) => (
          <button
            type="button"
            draggable
            onDragStart={(e) => setFieldDragPayload(e, v.field, "values")}
            onClick={() => {
              setPanel("menu");
              toggle();
            }}
            className="flex w-full cursor-grab items-center justify-between gap-1 rounded border border-slate-300 bg-white px-2 py-1.5 text-left text-xs font-medium text-ink shadow-sm hover:bg-slate-50 active:cursor-grabbing"
          >
            <span className="min-w-0 truncate">{label}</span>
            <ChevronDown className="h-3.5 w-3.5 shrink-0 text-muted" />
          </button>
        )}
      >
        {(close) =>
          panel === "menu" ? (
            // Matches real Excel's own Values-pill dropdown exactly: Move Up/Down/Beginning/End,
            // then Move to [other area], then Remove Field / Value Field Settings… — "Value Field
            // Settings…" just switches this SAME popover to the settings view below, rather than
            // opening a second one. Move Up/Down/etc. reuse the exact reorder handleValueRowDrop
            // already does via drag, just as discrete commands; disabled (not hidden) at either
            // end, same as Excel's own grayed-out state when there's nothing to move.
            <div className="flex flex-col gap-1 text-sm">
              {[
                { label: "Move Up", onSelect: onMoveUp, disabled: isFirst },
                { label: "Move Down", onSelect: onMoveDown, disabled: isLast },
                { label: "Move to Beginning", onSelect: onMoveToBeginning, disabled: isFirst },
                { label: "Move to End", onSelect: onMoveToEnd, disabled: isLast },
              ].map((item) => (
                <button
                  key={item.label}
                  type="button"
                  disabled={item.disabled}
                  onClick={
                    item.disabled
                      ? undefined
                      : () => {
                          item.onSelect();
                          close();
                        }
                  }
                  className={`rounded-lg px-2 py-1.5 text-left ${
                    item.disabled ? "cursor-not-allowed text-muted/50" : "hover:bg-gold/10"
                  }`}
                >
                  {item.label}
                </button>
              ))}
              <div className="my-1 border-t border-slate-100" />
              {[
                { label: "Move to Report Filter", onSelect: () => onMoveTo("filters") },
                { label: "Move to Row Labels", onSelect: () => onMoveTo("rows") },
                { label: "Move to Column Labels", onSelect: () => onMoveTo("columns") },
              ].map((item) => (
                <button
                  key={item.label}
                  type="button"
                  onClick={() => {
                    item.onSelect();
                    close();
                  }}
                  className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10"
                >
                  {item.label}
                </button>
              ))}
              <div className="my-1 border-t border-slate-100" />
              <button type="button" onClick={() => { onRemove(); close(); }} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
                Remove Field
              </button>
              <button type="button" onClick={() => setPanel("settings")} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
                Value Field Settings…
              </button>
              {/* Separate top-level entry, a sibling of "Value Field Settings…" rather than a tab
                  inside it — real Excel's Value Field Settings dialog only ever has Summarize
                  Values By / Show Values As / Number Format; Conditional Formatting is a wholly
                  separate feature reached its own way, not part of that dialog (moved out here
                  2026-10-05, flagged live as "زيادة عن الإكسل" when it lived inside Settings). */}
              <button type="button" onClick={() => setPanel("conditionalFormat")} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
                Conditional Formatting…
              </button>
            </div>
          ) : panel === "conditionalFormat" ? (
            <ConditionalFormatEditor
              format={v.conditionalFormat}
              onUpdate={(patch) => onUpdate({ conditionalFormat: { ...(v.conditionalFormat ?? { type: "None" }), ...patch } })}
              onBack={() => setPanel("menu")}
            />
          ) : (
            <ValueFieldSettingsPanel v={v} onUpdate={onUpdate} onUpdateFormat={onUpdateFormat} onBack={() => setPanel("menu")} rows={rows} />
          )
        }
      </Popover>
    </div>
  );
}

// Excel's real "Value Field Settings" dialog (Source Name/Custom Name, Summarize Values By / Show
// Values As tabs incl. the Base Field/Item picker, and the Number Format sub-panel) — factored out
// of ValuePill so PivotGrid's value-cell right-click context menu can open the exact same dialog
// (reached there via its own "Value Field Settings…" item) instead of duplicating this JSX a second
// time. Owns its own Number Format sub-navigation (`panel`) and tab state, independent of whatever
// panel-switching the CALLER uses around it (ValuePill's own "menu"/"settings"/"conditionalFormat",
// or PivotGrid's context menu's own root/showValuesAs) — `onBack` is the only thing the caller
// controls, fired when this dialog's own top-level Back button is pressed.
export function ValueFieldSettingsPanel({ v, onUpdate, onUpdateFormat, onBack, rows }) {
  const format = v.format ?? DEFAULT_FORMAT;
  // Excel's own Base Field/Base Item picker isn't PercentOf-only — Difference From and % Difference
  // From can ALSO target a specific named row instead of their default "(previous)" (extended
  // 2026-10-05, see MdxPivotQueryBuilder.ValidateShowValuesAs's own WHY comment). Only PercentOf
  // actually REQUIRES a choice here; for the other two, leaving Base Field unset keeps "(previous)".
  const showsBaseItemPicker = v.showValuesAs === "PercentOf" || v.showValuesAs === "DifferenceFrom" || v.showValuesAs === "PercentDifferenceFrom";
  const requiresBaseItem = v.showValuesAs === "PercentOf";
  const isDimensionField = v.field.startsWith("[");
  const dimensionAggregation = v.dimensionAggregation ?? "Count";
  const sourceName = isDimensionField ? formatDimensionName(v.field) : v.field;
  const computedLabel = isDimensionField ? `${dimensionAggregation} of ${sourceName}` : sourceName;
  const [panel, setPanel] = useState("settings");
  // Excel's real Value Field Settings dialog is two TABS ("Summarize Values By" / "Show Values
  // As"), not one flat stacked panel — flagged live, 2026-10-05, against screenshots of that real
  // dialog.
  const [settingsTab, setSettingsTab] = useState("summarize");
  const [baseItemOptions, setBaseItemOptions] = useState([]);
  const [loadingBaseItems, setLoadingBaseItems] = useState(false);
  // See FilterRow's own membersFailed WHY comment — same silent-failure gap, same fix.
  const [baseItemsFailed, setBaseItemsFailed] = useState(false);
  const [baseItemsRetryToken, setBaseItemsRetryToken] = useState(0);

  // Base item options come from the SAME member-search endpoint FilterRow uses — only fetched once
  // a Base field is actually chosen, and re-fetched whenever it changes.
  useEffect(() => {
    if (!showsBaseItemPicker || !v.baseField) {
      setBaseItemOptions([]);
      return;
    }
    let cancelled = false;
    const controller = new AbortController();
    setLoadingBaseItems(true);
    setBaseItemsFailed(false);
    getMembers(v.baseField, undefined, controller.signal)
      .then((members) => {
        if (!cancelled) setBaseItemOptions(members);
      })
      .catch((err) => {
        if (!cancelled && err.name !== "AbortError") setBaseItemsFailed(true);
      })
      .finally(() => {
        if (!cancelled) setLoadingBaseItems(false);
      });
    return () => {
      cancelled = true;
      controller.abort();
    };
  }, [showsBaseItemPicker, v.baseField, baseItemsRetryToken]);

  if (panel === "numberFormat") {
    return <NumberFormatEditor format={format} onUpdateFormat={onUpdateFormat} onBack={() => setPanel("settings")} />;
  }

  return (
    <div className="space-y-3">
      <button type="button" onClick={onBack} className="mb-1 inline-flex items-center gap-1 text-xs font-medium text-muted hover:text-ink">
        <ChevronLeft className="h-3.5 w-3.5" /> Back
      </button>
      {/* Matches Excel's own "Value Field Settings" dialog shape: Source Name (plain field,
          read-only) + Custom Name (editable, defaults to the computed "<Aggregation> of
          <field>" label) sit ABOVE the tab strip, same as the real dialog — flagged live,
          2026-10-04/05, against screenshots of that real dialog. */}
      <div className="space-y-1 rounded-md bg-slate-50 p-2">
        <div className="flex items-center gap-2 text-xs">
          <span className="font-medium text-muted">Source Name:</span>
          <span className="text-ink">{sourceName}</span>
        </div>
        <div>
          <span className="mb-1 block text-xs font-medium text-muted">Custom Name</span>
          <input
            type="text"
            className="input-field w-full py-1.5 text-sm"
            value={v.customLabel ?? ""}
            placeholder={computedLabel}
            onChange={(e) => onUpdate({ customLabel: e.target.value || null })}
          />
        </div>
      </div>

      <div className="flex gap-4 border-b border-slate-200">
        {[
          { key: "summarize", label: "Summarize Values By" },
          { key: "showAs", label: "Show Values As" },
        ].map((tab) => (
          <button
            key={tab.key}
            type="button"
            onClick={() => setSettingsTab(tab.key)}
            className={`-mb-px border-b-2 px-0.5 pb-1.5 text-xs font-medium ${
              settingsTab === tab.key ? "border-gold text-gold-strong" : "border-transparent text-muted hover:text-ink"
            }`}
          >
            {tab.label}
          </button>
        ))}
      </div>

      {settingsTab === "summarize" ? (
        isDimensionField ? (
          <div>
            <p className="mb-2 text-xs text-muted">Choose the type of calculation to use to summarize this field.</p>
            <select className={`${nativeSelectClass} w-full`} value={dimensionAggregation} onChange={(e) => onUpdate({ dimensionAggregation: e.target.value })}>
              {["Count", "Sum", "Average", "Max", "Min"].map((opt) => (
                <option key={opt} value={opt}>
                  {opt}
                </option>
              ))}
            </select>
          </div>
        ) : (
          // Real cube measures come with their aggregation pre-built into the cube itself —
          // confirmed live against the real cube (see PivotModels.PivotValueField's own WHY
          // comment) — there's no field-level value to re-aggregate here the way Excel's own
          // generic (non-OLAP) PivotTable can.
          <p className="text-xs text-muted">This is a real cube measure — its calculation is already built into the cube and can&apos;t be changed here.</p>
        )
      ) : (
        <div className="space-y-3">
          <select
            className={`${nativeSelectClass} w-full`}
            value={v.showValuesAs}
            onChange={(e) => {
              const next = e.target.value;
              // Both patches must land in ONE onUpdate call — two separate onChange calls would
              // each build off the same stale `v`/`value` closure, so the second would silently
              // discard the first (same race already caught and fixed in PivotGrid.jsx's own Show
              // Values As flyout, missed here when this panel was extracted out of ValuePill — fixed
              // now too, 2026-10-06). requireGrandTotals is a THIRD thing that needs the same
              // one-call treatment (it patches request.grandTotals, a level up from this value
              // field), so it rides along as a second argument instead of a separate onRequireGrandTotals
              // call — onUpdate's real implementations (FieldPicker.jsx's updateValue,
              // PivotGrid.jsx's own) fold it into the SAME onChange they already make.
              const patch = { showValuesAs: next };
              // Only auto-switch away from the untouched default — never overrides a format the
              // user already picked on purpose (e.g. they want a percent ratio shown as plain
              // Number for some reason).
              if (PERCENT_SHOW_VALUES_AS.has(next) && format.type === "General") {
                patch.format = { ...format, type: "Percentage" };
              }
              // % of Grand Total and Index both divide by the Grand Total — MdxPivotQueryBuilder
              // rejects them outright without it (see ValidateShowValuesAs), the same way real
              // Excel auto-enables Grand Totals the moment you pick either of these rather than
              // leaving you to hit a validation error yourself (flagged live, 2026-10-06, after
              // exactly that error surfaced with no indication of what to do about it).
              const requireGrandTotals = next === "PercentOfGrandTotal" || next === "Index";
              onUpdate(patch, { requireGrandTotals });
            }}
          >
            {SHOW_VALUES_AS_OPTIONS.map((opt) => (
              <option key={opt.value} value={opt.value}>
                {opt.label}
              </option>
            ))}
          </select>
          {showsBaseItemPicker && (
            <div className="space-y-2 rounded-md bg-slate-50 p-2">
              <div>
                <span className="mb-1 block text-xs font-medium text-muted">Base field (Rows)</span>
                {/* Same no-width-class bug as the Symbol select — a real field's display name has no
                    length limit. */}
                <select className={`${nativeSelectClass} w-full`} value={v.baseField ?? ""} onChange={(e) => onUpdate({ baseField: e.target.value || null, baseItem: "" })}>
                  {/* Difference From/% Difference From can leave this as their own default
                      "(previous)" base — PercentOf has no such fallback, so it must actually
                      pick a field. */}
                  <option value="">{requiresBaseItem ? "Choose a row field…" : "(previous)"}</option>
                  {rows.map((field) => (
                    <option key={field} value={field}>
                      {formatDimensionName(field)}
                    </option>
                  ))}
                </select>
              </div>
              {v.baseField && baseItemsFailed && (
                <div className="flex items-center justify-between gap-2 rounded-xl border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">
                  <span>Couldn't load values — the cube may be busy.</span>
                  <button
                    type="button"
                    onClick={() => setBaseItemsRetryToken((t) => t + 1)}
                    className="shrink-0 rounded-lg border border-rose-300 px-2 py-1 text-xs font-medium hover:bg-rose-100"
                  >
                    Retry
                  </button>
                </div>
              )}
              {v.baseField && !baseItemsFailed && (
                <div>
                  <span className="mb-1 block text-xs font-medium text-muted">Base item</span>
                  {/* Same no-width-class bug — a real member's own label has no length limit either. */}
                  <select className={`${nativeSelectClass} w-full`} value={v.baseItem ?? ""} onChange={(e) => onUpdate({ baseItem: e.target.value })} disabled={loadingBaseItems}>
                    <option value="">{loadingBaseItems ? "Loading…" : "Choose a value…"}</option>
                    {baseItemOptions.map((member) => (
                      <option key={member} value={member}>
                        {member}
                      </option>
                    ))}
                  </select>
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {/* Excel's own "Number Format…" button, sitting below the tab strip (not inside either tab)
          — opens the SEPARATE Format Cells dialog, same relationship the real Value Field Settings
          dialog has with it. */}
      <div className="flex items-center justify-between gap-2 rounded-md bg-slate-50 p-2">
        <span className="text-xs text-muted">
          Format: <span className="font-medium text-ink">{format.type}</span>
        </span>
        <button type="button" onClick={() => setPanel("numberFormat")} className="btn-secondary px-2 py-1 text-xs">
          Number Format…
        </button>
      </div>
    </div>
  );
}

export default function FieldPicker({
  value,
  onChange,
  dimensions,
  measures,
  tables,
  tablesLoading,
  enabledTableNames,
  onEnabledTableNamesChange,
  fieldsLoading,
  onReset,
  // Real Excel's per-field "Field Settings…" > Custom Name / Layout & Print (PLAN 6 Phase 2) — see
  // PivotBuilder.jsx's own fieldLayouts state WHY comment for why this is a sibling pair of props,
  // not part of value/onChange.
  fieldLayouts,
  onFieldLayoutsChange,
  // Report-wide layout + Repeat All Item Labels, so Field Settings can show a field's EFFECTIVE
  // Compact/Repeat state (what it inherits when it has no override of its own), like Excel does.
  reportLayout = "Compact",
  repeatItemLabels = false,
}) {
  // Measures the Rows/Columns/Values/Filters zones column's OWN rendered height, so the Fields tree
  // card next to it can match it exactly in EITHER direction (requested live 2026-10-12, "خلي الـ
  // fields تاخد المكان الفاضي" — after a plain CSS grid `items-stretch` attempt made things WORSE,
  // not better: stretch can only GROW a shorter item to match a taller one, never shrink a naturally
  // huge one down, so it just stretched the short zones column to match the field list's own full,
  // uncapped ~300-table height instead of the other way around). A ResizeObserver on this ONE
  // sibling element is a much narrower, safer measurement than the "rest of the viewport" approach
  // that caused a real page-scrollbar bug the last time dynamic sizing was tried here (2026-10-05) —
  // this only ever reacts to the zones column's own height, never the page's.
  const zonesColumnRef = useRef(null);
  const [zonesColumnHeight, setZonesColumnHeight] = useState(null);
  useEffect(() => {
    const el = zonesColumnRef.current;
    if (!el) return undefined;
    const measure = () => setZonesColumnHeight(el.getBoundingClientRect().height);
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

  const dimensionOptions = dimensions.map((d) => ({ value: d.field, label: d.displayName }));
  const measureOptions = measures.map((m) => ({ value: m.field, label: m.displayName }));
  // A calculated field's own LeftField/RightField must reference a real cube measure only (no
  // chaining calculated-on-calculated — see the backend's CalculatedField), but once FULLY
  // declared (name AND both measures — matches PivotBuilder.jsx's own completeness check before
  // sending a query) its name becomes usable anywhere a regular measure is, including as a Value.
  // Requiring only `f.name` here used to let an incomplete calculated field be picked in Values —
  // a real bug caught live 2026-09-27: the backend would then try to reference a measure that was
  // never actually declared, since PivotBuilder.jsx drops an incomplete calculated field from the
  // query entirely while the dangling Value reference stayed selected.
  const calculatedFieldOptions = value.calculatedFields
    .filter((f) => f.name && f.leftField && f.rightField)
    .map((f) => ({ value: f.name, label: `${f.name} (calculated)` }));
  const valueFieldOptions = [...measureOptions, ...calculatedFieldOptions];

  // `options.requireGrandTotals` rides along in the SAME onChange call as the value patch — see
  // ValueFieldSettingsPanel's own Show Values As handler for why this can't be two separate calls
  // (PercentOfGrandTotal/Index both need request.grandTotals turned on, a level up from this
  // one value field's own patch). Forces the safe universal superset (RowsAndColumns) rather than
  // the specific axis this variant actually needs — see PivotGrid.jsx's own identical comment.
  const updateValue = (index, patch, options = {}) => {
    const copy = [...value.values];
    copy[index] = { ...copy[index], ...patch };
    onChange({ ...value, values: copy, ...(options.requireGrandTotals ? { grandTotals: "RowsAndColumns" } : {}) });
  };

  const updateValueFormat = (index, patch) => {
    const current = value.values[index].format ?? DEFAULT_FORMAT;
    updateValue(index, { format: { ...current, ...patch } });
  };

  const removeValue = (index) => {
    const values = value.values.filter((_, i) => i !== index);
    onChange(clearDanglingMeasureReferences({ ...value, values }, values.map((v) => v.field)));
  };

  // Backs ValuePill's right-click "Move Up"/"Move Down"/"Move to Beginning"/"Move to End" — the
  // same reorder handleValueRowDrop already does via drag, just as discrete menu actions (Excel's
  // own context menu offers both the drag AND these explicit commands for the same reorder).
  const moveValueTo = (fromIndex, toIndex) => {
    if (toIndex < 0 || toIndex >= value.values.length || toIndex === fromIndex) return;
    const next = [...value.values];
    const [item] = next.splice(fromIndex, 1);
    next.splice(toIndex, 0, item);
    onChange({ ...value, values: next });
  };

  // The single add/remove entry point for Values, driven by the multi-select above the pill list
  // (see the "Values" box below) — a plain diff against the current pills: fields no longer
  // selected are dropped, newly-selected fields become new pills, and anything unchanged KEEPS
  // its existing config (showValuesAs/format/conditionalFormat) rather than resetting it, since
  // re-selecting the same measure isn't meant to wipe out per-measure settings already made.
  const setSelectedValueFields = (selectedFields) => {
    // A dimension field dropped directly into Values (Excel's own "Count of <field>" default for a
    // non-measure field) is managed entirely by drag, not by this dropdown — this control's own
    // `options` only ever lists real measures/calculated fields, so `selectedFields` can never
    // contain a dimension path. Without separating them out first, every edit here (even
    // unselecting ONE unrelated real measure) would silently wipe out every dimension-based value,
    // since `selectedFields.includes(v.field)` is always false for a path this dropdown never
    // offered in the first place (caught live, 2026-10-05).
    const dimensionBased = value.values.filter((v) => v.field.startsWith("["));
    const measureBased = value.values.filter((v) => !v.field.startsWith("["));
    const kept = measureBased.filter((v) => selectedFields.includes(v.field));
    const added = selectedFields
      .filter((field) => !measureBased.some((v) => v.field === field))
      .map((field) => ({ field, showValuesAs: "Normal" }));
    const values = [...dimensionBased, ...kept, ...added];
    onChange(clearDanglingMeasureReferences({ ...value, values }, values.map((v) => v.field)));
  };

  const updateFilter = (index, patch) => {
    const copy = [...value.filters];
    copy[index] = patch;
    onChange({ ...value, filters: copy });
  };

  const removeFilter = (index) => {
    onChange({ ...value, filters: value.filters.filter((_, i) => i !== index) });
  };

  // Real Excel never shows a "choose a field" control inside a filter's own dropdown — the field
  // is always already fixed by the time you open one (see FilterRow's own WHY comment). "+" used to
  // add a genuinely EMPTY filter needing FilterRow's own inline picker to fix that; now picks the
  // field FIRST, in this small popover (mirrors the Values zone's own "+" popover, just a single
  // Select instead of isMulti — one filter added per click, not a bulk add), so FilterRow always
  // receives an already-complete filter and never needs its inline picker at all (caught live,
  // 2026-10-12, "ليه عامل search , drop down مع ان في الexcel واحدة بس").
  const addFilter = (field) => {
    if (!field || value.filters.some((f) => f.field === field)) return;
    onChange({ ...value, filters: [...value.filters, { field, mode: "Members", includedMembers: [] }] });
  };

  // Excel's classic "drag a field between zones" — see HierarchicalCubePathSelect/ZoneDropArea
  // for the native-HTML5-DnD half of this; this is the other half, since only FieldPicker holds
  // every zone's array at once.
  //
  // Rows/Columns/Values are MUTUALLY EXCLUSIVE (a field moving into one leaves the other two,
  // same as Excel — a field can't be a row, a column, and a value all at once). Filters is
  // ADDITIVE instead: dropping a field onto Filters adds a member-filter for it WITHOUT removing
  // it from wherever else it already is, and dragging a Filters pill onto Rows/Columns/Values
  // just adds it there too, leaving the filter in place — this mirrors a real feature this tool
  // already had before drag-and-drop existed (a Row/Column field carrying its own member filter,
  // same as Excel's per-column filter icon, distinct from the classic Report Filter area), so
  // drag-and-drop had to preserve it rather than force strict one-zone-only placement.
  const moveField = (field, fromZone, toZone) => {
    if (fromZone === toZone) {
      return;
    }

    const next = { ...value };

    // "available" = drag a placed field back out to remove it, Excel's own "drag out of any area
    // to remove" — just strip it from everywhere (requested live, 2026-09-28, alongside each
    // pill's existing X button, not instead of it). Shares removeFieldEverywhere with
    // PivotBuilder.jsx's own page-level "dropped outside every zone" fallback — this is the
    // "dropped specifically back on the Available Fields panel" case of that same removal.
    if (toZone === "available") {
      onChange(removeFieldEverywhere(value, field));
      return;
    }

    // Real Excel never lets one field sit in more than one of Report Filter / Row Labels / Column
    // Labels at the same time — dropping it into one of those three always pulls it out of
    // whichever of the other two it was in (confirmed live, 2026-10-05: "نمنع انه يبقى في
    // مكانين"). Values is a DIFFERENT kind of area (an aggregation, not a placement) and stays
    // fully independent — Excel routinely has the same field in both Rows and Values at once
    // (e.g. "Doctor Name" in Rows + "Count of Doctor Name" in Values), so toZone === "values"
    // deliberately never touches rows/columns/filters, and landing in rows/columns/filters never
    // touches values either.
    if (toZone === "rows" || toZone === "columns" || toZone === "filters") {
      // rowLevelSorts'/columnLevelSorts' own `level` values are positional indexes into
      // Rows/Columns — this function can both REMOVE the field from one of those axes (it was
      // already there, now moving elsewhere) and APPEND it to one (toZone is that axis), either of
      // which shifts what index every OTHER level-sort entry actually points to. The
      // HierarchicalCubePathSelect onChange handlers below (same-zone reorder, or the "+" add
      // popover) already clear these on any add/remove/reorder — this cross-zone move path (plain
      // drag-and-drop between zones, and each pill's own right-click "Move to…" menu) never did,
      // leaving a stale level-sort behind that the backend would reject on the very next query
      // (caught 2026-10-06, full codebase review).
      const rowsAffected = value.rows.includes(field) || toZone === "rows";
      const columnsAffected = value.columns.includes(field) || toZone === "columns";
      next.rows = value.rows.filter((f) => f !== field);
      next.columns = value.columns.filter((f) => f !== field);
      next.filters = value.filters.filter((f) => f.field !== field);
      if (rowsAffected) {
        next.rowLevelSorts = [];
      }
      if (columnsAffected) {
        next.columnLevelSorts = [];
      }
      // A field leaving Rows takes its grouping with it — the backend rejects a grouping on a
      // field that's no longer in Rows (ValidateGroupings).
      next.dateGroupings = value.dateGroupings.filter((g) => g.field !== field);
      next.numericGroupings = value.numericGroupings.filter((g) => g.field !== field);
      // Same orphaned-per-field-config cleanup, for the Field Settings Subtotals override
      // (2026-10-13) — meaningless once the field it was configured for has left Rows.
      next.fieldSubtotals = value.fieldSubtotals.filter((f) => f.field !== field);
      next.fieldsShowingItemsWithNoData = value.fieldsShowingItemsWithNoData.filter((f) => f !== field);
    }
    // Values deliberately does NOT strip/replace an existing entry for this same field here (unlike
    // rows/columns/filters above, which a field can only ever occupy once) — dragging an
    // already-placed measure into Values AGAIN is exactly how real Excel lets you add the SAME
    // measure a second time (e.g. once Normal, once % of Grand Total — a completely ordinary
    // pattern, and the backend already accepts it, see MdxPivotQueryBuilder.ValidateDuplicateValues).
    // Previously this branch stripped the existing entry then re-added a fresh default one in its
    // place, so re-dragging a measure already in Values just silently RESET it instead of adding a
    // second copy — there was actually no way at all to get two entries for the same field (caught
    // 2026-10-06, "ازاي اضيف ال measure اكتر من مرة" — the user went looking for this and genuinely
    // couldn't find it, because it didn't exist yet).

    if (toZone === "rows" && !next.rows.includes(field)) {
      next.rows = [...next.rows, field];
    } else if (toZone === "columns" && !next.columns.includes(field)) {
      next.columns = [...next.columns, field];
    } else if (toZone === "values") {
      // Real Excel's own default the moment a non-measure field is dropped into Values: Sum for a
      // NUMERIC field, Count for anything else (confirmed live, 2026-10-05 — a real "Age" field
      // defaulted to "Sum of Age" in Excel, not "Count of Age", which is all this app ever
      // defaulted to before). A real cube measure (not in `dimensions` at all) never takes this
      // path's dimensionAggregation — BuildMeasureExpression ignores it entirely for those.
      const isNumericDimension = dimensions.some((d) => d.field === field && d.isNumeric);
      const newEntry = { field, showValuesAs: "Normal", dimensionAggregation: isNumericDimension ? "Sum" : "Count" };
      // A SECOND entry for the same field needs its OWN distinct caption — every column-to-measure
      // lookup in this app (buildCaptionToIndex on the frontend, PivotModels.BuildCaptionToIndex on
      // the backend, used by GrandTotalsProcessor/SubtotalsProcessor/ShowValuesAsProcessor/
      // NumberFormatter/the live grid's own rendering) is keyed by the measure's own CAPTION, on the
      // documented assumption that every Value has a unique one — true until this exact "add the
      // same measure twice" feature existed. Two unlabeled "Total Revenue" entries sharing the
      // identical caption broke that assumption: the lookup could only ever find ONE of them, so the
      // second silently matched whatever ELSE happened to be the next distinct caption in the actual
      // header data (caught live, 2026-10-06, right after enabling duplicates — the second "Total
      // Revenue" column showed a stray extra all-zero "Total" column instead of real numbers,
      // "هو كدا صح؟"). Auto-naming the duplicate here (same idea as real Excel's own "Sum of
      // Revenue2") sidesteps rewriting that caption-matching algorithm everywhere it's relied
      // upon — it just keeps the one invariant the whole pipeline already assumes always true.
      if (value.values.some((v) => v.field === field)) {
        newEntry.customLabel = nextAvailableDuplicateLabel(field, value.values);
      }
      next.values = [...next.values, newEntry];
    } else if (toZone === "filters" && !next.filters.some((f) => f.field === field)) {
      next.filters = [...next.filters, { field, mode: "Members", includedMembers: [] }];
    }

    next.calculatedItems = clearOrphanedCalculatedItems(value.calculatedItems, next.rows, next.columns);
    // Same dangling-reference cleanup removeValue/setSelectedValueFields and the Rows onChange
    // handler already apply — this is the THIRD path that can remove a field from Values or Rows
    // (plain drag-and-drop between zones, and each pill's own right-click "Move to…" menu).
    next.values = clearDanglingBaseFieldReferences(next.values, next.rows);
    Object.assign(next, clearDanglingMeasureReferences(next, next.values.map((v) => v.field)));
    // Drilldown v3: the drilldown is a Rows slot now — it survives this move unless its own slot
    // left Rows or a same-hierarchy field arrived (see reconcileDrilldowns — it can also follow its pill to Columns).
    Object.assign(next, reconcileDrilldowns(next));

    onChange(next);
  };

  // Values order matters (it decides column grouping order, same as Rows/Columns — see the
  // "measures are the fastest-varying factor" invariant in MdxPivotQueryBuilder), so — unlike
  // Filters, which is an unordered set of conditions — a Values pill needs the same
  // drag-to-reorder-within-the-zone support HierarchicalCubePathSelect already gives Rows/Columns.
  // stopPropagation keeps ZoneDropArea's own onDrop (cross-zone move, appended at the end) from
  // also firing for the same event.
  const handleValueRowDrop = (e, targetIndex) => {
    e.preventDefault();
    e.stopPropagation();
    const payload = readFieldDragPayload(e);
    if (!payload) return;
    if (payload.source !== "values") {
      moveField(payload.field, payload.source, "values");
      return;
    }
    const draggedIndex = value.values.findIndex((v) => v.field === payload.field);
    if (draggedIndex === -1 || draggedIndex === targetIndex) return;
    const next = [...value.values];
    const [item] = next.splice(draggedIndex, 1);
    const insertAt = draggedIndex < targetIndex ? targetIndex - 1 : targetIndex;
    next.splice(insertAt, 0, item);
    onChange({ ...value, values: next });
  };

  // Filters are logically an unordered set of AND-conditions (query result is identical regardless
  // of order, see moveField's own WHY comment) — this is purely so the pills can be dragged into
  // whatever visual order the user finds clearer, matching every other zone's drag-reorder support
  // (Rows/Columns via HierarchicalCubePathSelect, Values via handleValueRowDrop above) rather than
  // being the one zone where a field pill just can't be reordered (caught live, 2026-09-28). Same
  // splice-by-index shape as handleValueRowDrop.
  const handleFilterRowDrop = (e, targetIndex) => {
    e.preventDefault();
    e.stopPropagation();
    const payload = readFieldDragPayload(e);
    if (!payload) return;
    if (payload.source !== "filters") {
      moveField(payload.field, payload.source, "filters");
      return;
    }
    const draggedIndex = value.filters.findIndex((f) => f.field === payload.field);
    if (draggedIndex === -1 || draggedIndex === targetIndex) return;
    const next = [...value.filters];
    const [item] = next.splice(draggedIndex, 1);
    const insertAt = draggedIndex < targetIndex ? targetIndex - 1 : targetIndex;
    next.splice(insertAt, 0, item);
    onChange({ ...value, filters: next });
  };

  // Adds EVERY level of a multi-level hierarchy to Rows/Columns in ONE action, in the hierarchy's
  // own top-to-bottom level order — OLAP +/- drill-down Phase 1 (2026-10-07, see
  // EXCEL_PARITY_AUDIT.md): until real per-member expand/collapse exists (Phase 2, a much larger
  // architectural change), this at least removes the need to hunt for and add a hierarchy's levels
  // one at a time, with their correct nesting order guaranteed. Mirrors moveField's own
  // rows/columns placement rules (same-field exclusivity across Filters/Rows/Columns, level-sort/
  // grouping cleanup, dangling-reference cleanup) but for N fields landing in one onChange instead
  // of moveField's one — calling moveField in a loop here would silently drop all but the last
  // field, since each call closes over the same pre-update `value` from this render.
  const addHierarchyToAxis = (levelFields, axis) => {
    const otherAxis = axis === "rows" ? "columns" : "rows";
    const fieldsToAdd = levelFields.filter((f) => !value[axis].includes(f));
    if (fieldsToAdd.length === 0) return;

    const next = { ...value };
    next[axis] = [...value[axis], ...fieldsToAdd];
    next[otherAxis] = value[otherAxis].filter((f) => !levelFields.includes(f));
    next.filters = value.filters.filter((f) => !levelFields.includes(f.field));
    next.rowLevelSorts = [];
    next.columnLevelSorts = [];
    next.dateGroupings = value.dateGroupings.filter((g) => !levelFields.includes(g.field));
    next.numericGroupings = value.numericGroupings.filter((g) => !levelFields.includes(g.field));
    // Same orphaned-per-field-config cleanup as moveField's own (2026-10-13) — only matters when
    // axis is "columns" (levelFields leaving Rows), harmless no-op otherwise since none of
    // levelFields would have a fieldSubtotals entry to begin with if they were already in Columns.
    next.fieldSubtotals = value.fieldSubtotals.filter((f) => !levelFields.includes(f.field));
    next.fieldsShowingItemsWithNoData = value.fieldsShowingItemsWithNoData.filter((f) => !levelFields.includes(f));
    next.calculatedItems = clearOrphanedCalculatedItems(value.calculatedItems, next.rows, next.columns);
    next.values = clearDanglingBaseFieldReferences(next.values, next.rows);
    Object.assign(next, clearDanglingMeasureReferences(next, next.values.map((v) => v.field)));
    // Drilldown v3: same rule as moveField's — "+ Rows" ends the drilldown only if it adds a level
    // of the drilldown's OWN hierarchy (or "+ Cols" takes its slot away).
    Object.assign(next, reconcileDrilldowns(next));

    onChange(next);
  };

  // Real Excel's per-member OLAP drilldown (PLAN 8). Since Drilldown v3 (2026-10-14) the hierarchy
  // JOINS Rows as one slot (its top-level field, appended at the end like any field added to Rows)
  // instead of replacing Rows — the other Rows fields stay, and the slot can be dragged/reordered
  // among them like a normal pill (see reconcileDrilldowns for when drilldown mode ends). Any flat
  // levels of this SAME hierarchy already in Rows are dropped: SSAS can't put one hierarchy on an
  // axis twice. Σ Values moves to Columns (the backend requires it there for a combined drilldown),
  // and show-items-with-no-data is cleared (not supported alongside a combined drilldown).
  const startDrilldown = (topLevelField) => {
    const sameHierarchy = (f) => f.split("].[").slice(0, 2).join("].[") === topLevelField.split("].[").slice(0, 2).join("].[");
    const otherRows = value.rows.filter((f) => !sameHierarchy(f) && f !== MEASURES_PLACEHOLDER);
    // The same hierarchy can't sit on Columns too (flat levels or a Columns drilldown).
    const otherColumns = value.columns.filter((f) => !sameHierarchy(f));
    const columns =
      value.rows.includes(MEASURES_PLACEHOLDER) && !otherColumns.includes(MEASURES_PLACEHOLDER) ? [...otherColumns, MEASURES_PLACEHOLDER] : otherColumns;
    const columnsDrill = value.columnsHierarchyDrilldown;
    onChange({
      ...value,
      rows: [...otherRows, topLevelField],
      columns,
      rowsHierarchyDrilldown: { topLevelField, expandedMemberUniqueNames: [] },
      columnsHierarchyDrilldown: columnsDrill && !sameHierarchy(columnsDrill.topLevelField) ? columnsDrill : null,
      rowLevelSorts: [],
      fieldSubtotals: value.fieldSubtotals.filter(
        (s) => otherRows.includes(s.field) || (columnsDrill && !sameHierarchy(columnsDrill.topLevelField) && s.field === columnsDrill.topLevelField),
      ),
      fieldsShowingItemsWithNoData: [],
      // Drilldown v2 (2026-10-14): Subtotals now applies to the hierarchy, and Excel's own default
      // for a new hierarchy is subtotals at Top (each expanded member's row shows its total). This
      // app's report default is Off, which would blank every expanded member — so a fresh
      // drilldown starts at Top; any explicit choice made afterward is kept.
      subtotals: value.subtotals === "Off" ? "Top" : value.subtotals,
      // A per-level sort can't apply to a ragged drilldown axis (MdxPivotQueryBuilder rejects it),
      // and a Rows `sort` aimed at the old flat fields would now sort the hierarchy instead.
      sort: value.sort?.axis === "Rows" ? null : value.sort,
    });
  };

  const isFieldPlaced = (field) =>
    value.rows.includes(field) || value.columns.includes(field) || value.values.some((v) => v.field === field) || value.filters.some((f) => f.field === field);

  return (
    // Field list on the LEFT, the 4 drop zones stacked in ONE column on the RIGHT — real Excel's
    // own PivotTable Fields pane layout (confirmed live, 2026-10-04, against a screenshot: "الاربعه
    // تحت بعض وعلى شمالهم الجدول اللي اختارته"). Replaces an earlier "2×2 grid" version that was
    // itself a correction made against a DIFFERENT real-Excel screenshot shown earlier — Excel's
    // pane apparently reflows between a single stacked column and a 2×2 grid depending on its own
    // width/version, so this now matches the most recent, most explicit screenshot+instruction
    // rather than re-arguing which one is "more canonical".
    // items-start: the field list card's HEIGHT is now driven explicitly by zonesColumnHeight (a
    // narrow, self-contained ResizeObserver measurement of the zones column alone — see this
    // function's own top-of-component WHY comment), not by grid's own align-items:stretch — stretch
    // was tried live 2026-10-12 ("خلي ارتفاع ال fields زي ال 4 اللي جنبه") and made things WORSE: it
    // can only GROW a shorter item to match a taller one's natural content size, never SHRINK a
    // naturally-huge one (this field list, ~300+ tables, uncapped) down — so it stretched the SHORT
    // zones column to match the field list's own enormous natural height instead ("انت شايف ان دا
    // كدا صح؟"), the opposite of intended. items-start here just means "don't ALSO let grid stretch
    // interfere" — the actual matching happens via the explicit pixel height set below instead.
    // minmax(0, Nfr), not just Nfr: a plain fr track defaults to a minimum of its own content's
    // min-content width — the classic CSS Grid version of the flexbox min-width:auto trap (see
    // PivotBuilder.jsx's own WHY comment on that one) — so a long unbreakable string anywhere
    // inside either column (e.g. a field name) was forcing that whole column, and the grid itself,
    // wider than the card actually sizes to, spilling a sliver of content out past the card's own
    // right edge (flagged live, 2026-10-04: "ليه طالع برا الكارد"). minmax(0, …) lets each column
    // actually shrink to the space it's given, so `truncate` inside it can do its job.
    <div className="grid grid-cols-[minmax(0,1.3fr)_minmax(0,0.7fr)] items-start gap-2">
      {/* Height set explicitly from the zones column's OWN measured height (zonesColumnHeight,
          above) once known — min 200px as a sane floor for the very first render (before the
          ResizeObserver's first callback fires) and for the unlikely case of an almost-empty zones
          column. AvailableFieldsPanel fills whatever height this div ends up with (h-full flex-col,
          its own scrollable list as flex-1) — see its own WHY comment. */}
      <div style={{ height: zonesColumnHeight ? `${Math.max(zonesColumnHeight, 200)}px` : undefined }}>
        <AvailableFieldsPanel
          dimensions={dimensions}
          measures={measures}
          valuesCount={value.values.length}
          isPlaced={isFieldPlaced}
          tables={tables}
          tablesLoading={tablesLoading}
          enabledTableNames={enabledTableNames}
          onEnabledTableNamesChange={onEnabledTableNamesChange}
          fieldsLoading={fieldsLoading}
          onFieldDropped={moveField}
          onAddHierarchy={addHierarchyToAxis}
          onStartDrilldown={startDrilldown}
        />
      </div>

      {/* Each zone shows only its pill list; the full add/edit form (dimension picker, filter
          mode, show-values-as, ...) lives in a Popover behind the zone's "+" or, for Filters/
          Values, behind the pill itself — one click away instead of permanently occupying space.
          ref={zonesColumnRef}: this is the sibling FieldPicker's own Fields card measures itself
          against — see zonesColumnHeight's own WHY comment above.
          pt-2, not the original pt-6 (cut 2026-10-12 alongside the zone boxes above, same reason):
          this existed purely so "Reset" lines up with "FIELDS" (.card's own p-6 top padding) —
          real, but strictly cosmetic, and the page-scrollbar problem took priority over it being
          pixel-perfect. Still a little breathing room instead of "Reset" flush against the edge. */}
      <div ref={zonesColumnRef} className="pt-2">
        {/* Real Excel's own instructional text used to sit here ("Drag fields between areas
            below:", confirmed live 2026-10-05 against a real Excel screenshot) — removed again
            2026-10-12 on explicit request ("شيل دا"): this sidebar is tight on vertical space now
            that the save/export controls also live in it (see PivotBuilder.jsx's own WHY comment),
            so literal Excel-instructional-text parity lost out to reclaiming the row for it. */}
        <div className="mb-1 flex justify-end">
          <button type="button" onClick={onReset} className="btn-secondary px-3 py-1.5 text-xs">
            Reset
          </button>
        </div>
        <div className="grid grid-cols-1 gap-1.5">
          <div className="flex flex-col rounded-xl border border-slate-200 bg-white p-2 shadow-sm">
            <div className="mb-1.5 flex items-center justify-between gap-2">
              <span className="flex items-center gap-1.5 text-xs font-bold uppercase tracking-wide text-muted">
                <Filter className="h-3.5 w-3.5" />
                Filters {value.filters.length > 0 && <span className="text-gold">({value.filters.length})</span>}
              </span>
              <Popover
                align="right"
                trigger={(toggle) => (
                  <button
                    type="button"
                    onClick={toggle}
                    className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-gold/15 text-gold hover:bg-gold/25"
                    aria-label="Add filter"
                    title="Add filter"
                  >
                    <Plus className="h-3.5 w-3.5" />
                  </button>
                )}
              >
                {(close) => (
                  <div>
                    <p className="mb-2 text-xs text-muted">Choose a field to filter.</p>
                    <Select
                      className="react-select-container"
                      classNamePrefix="react-select"
                      placeholder="Choose a field to filter…"
                      options={dimensionOptions}
                      value={null}
                      onChange={(opt) => {
                        addFilter(opt?.value);
                        close();
                      }}
                      menuPortalTarget={typeof document !== "undefined" ? document.body : null}
                      menuPosition="fixed"
                      styles={{ menuPortal: (base) => ({ ...base, zIndex: 10000 }) }}
                    />
                  </div>
                )}
              </Popover>
            </div>
            <ZoneDropArea
              zoneKey="filters"
              onFieldDropped={moveField}
              // Rejects a measure dropped here (e.g. dragged from a Values pill) — Filters is a
              // dimension-only zone; without this a dropped measure landed as a FilterRow whose
              // field isn't in dimensionOptions, showing no match in its own Select while the
              // broken value stayed set (caught live, 2026-09-28).
              isFieldAccepted={(field) => dimensionOptions.some((opt) => opt.value === field)}
              isEmpty={value.filters.length === 0}
            >
              <div className="flex flex-wrap gap-1.5">
                {value.filters.map((filter, index) => (
                  <FilterPill
                    key={index}
                    filter={filter}
                    dimensionOptions={dimensionOptions}
                    // TopN/BottomN's ByMeasureField must match a field already in Values (see
                    // MdxPivotQueryBuilder.ValidateFilters) — not just any cube measure, so this is
                    // value.values, not the full measureOptions list used elsewhere in this component.
                    selectedValueOptions={value.values.filter((v) => v.field).map((v) => ({ value: v.field, label: v.field }))}
                    onChange={(patch) => updateFilter(index, patch)}
                    onRemove={() => removeFilter(index)}
                    onDropReorder={(e) => handleFilterRowDrop(e, index)}
                    onMoveTo={(toZone) => moveField(filter.field, "filters", toZone)}
                  />
                ))}
              </div>
            </ZoneDropArea>
          </div>

          <div>
            <HierarchicalCubePathSelect
              // The field pane now sits docked at the page's right edge (see PivotBuilder.jsx) —
              // every floating Popover panel in here defaults to growing rightward off the edge
              // unless told to grow leftward instead, same fix already needed once for Rows/
              // Values before this layout change (2026-09-27).
              align="right"
              label="Columns"
              icon={Columns3}
              options={dimensions.map((d) => d.field)}
              value={value.columns}
              onChange={(columns) => {
                // moveField's own "remove from every destination-capable zone first, then add to
                // the target" logic only runs for drag-and-drop (onFieldDropped below) — THIS
                // onChange also fires from the "+" Add popover's own internal newPathsToAdd, which
                // only ever filtered against Columns' OWN existing value, with no visibility into
                // Rows at all. A field picked here that's already in Rows landed in BOTH axes at
                // once (crossjoining a field with itself in the generated MDX) — the one real
                // exclusivity gap the comment below used to claim didn't exist (caught 2026-10-06,
                // full codebase review). Stripped out of Rows here instead, same "last one wins"
                // rule moveField already enforces for every other path.
                const newlyAddedToColumns = columns.filter((f) => !value.columns.includes(f));
                const rows = value.rows.filter((f) => !newlyAddedToColumns.includes(f));
                // Dropping a column field also drops any grouping/calculated item that referenced
                // it — the backend rejects both on a field that's no longer in Rows/Columns
                // (ValidateGroupings, ValidateCalculatedItems). A field still present in Rows keeps
                // its grouping (it didn't actually leave either axis).
                const stillGrouped = (field) => columns.includes(field) || rows.includes(field);
                const dateGroupings = value.dateGroupings.filter((g) => stillGrouped(g.field));
                const numericGroupings = value.numericGroupings.filter((g) => stillGrouped(g.field));
                const calculatedItems = clearOrphanedCalculatedItems(value.calculatedItems, rows, columns);
                // columnLevelSorts' `level` values are positional indexes into Columns — any add,
                // remove, or reorder here can shift which field a stored index actually points to,
                // so they're cleared rather than risking a rule silently applying to the wrong
                // field (same defensive-clear approach as groupings/calculatedItems above).
                // Columns drilldown (C4): its slot rule, same as Rows' (see reconcileDrilldowns).
                const next = {
                  ...value,
                  rows,
                  columns,
                  dateGroupings,
                  numericGroupings,
                  calculatedItems,
                  columnLevelSorts: [],
                  rowLevelSorts: newlyAddedToColumns.length > 0 ? [] : value.rowLevelSorts,
                };
                onChange({ ...next, ...reconcileDrilldowns(next) });
              }}
              dragSourceKey="columns"
              onFieldDropped={moveField}
              isFieldAccepted={(field) => field === MEASURES_PLACEHOLDER || dimensionOptions.some((opt) => opt.value === field)}
              contextMenuTargets={[
                { zoneKey: "filters", label: "Move to Report Filter" },
                { zoneKey: "rows", label: "Move to Row Labels" },
                { zoneKey: "values", label: "Move to Values" },
              ]}
              // Columns drilldown (2026-10-14): its own Field Settings > Subtotals (Automatic/None)
              // decides the "Egypt Total" columns independently of the Rows subtotals — stored in the
              // same fieldSubtotals list (see ColumnDrilldownLayoutProcessor). Only that one field.
              fieldSettings={
                value.columnsHierarchyDrilldown
                  ? {
                      reportDefault: value.subtotals,
                      overrides: value.fieldSubtotals,
                      appliesTo: (field) => field === value.columnsHierarchyDrilldown.topLevelField,
                      noCustom: true,
                      onChange: (field, patch) => {
                        const existing = value.fieldSubtotals.find((f) => f.field === field);
                        const position = patch.position !== undefined ? patch.position : (existing?.position ?? null);
                        const withoutField = value.fieldSubtotals.filter((f) => f.field !== field);
                        onChange({ ...value, fieldSubtotals: position ? [...withoutField, { field, position, functions: null }] : withoutField });
                      },
                    }
                  : undefined
              }
            />
          </div>

          <div>
            <HierarchicalCubePathSelect
              align="right"
              label="Rows"
              icon={Rows3}
              options={dimensions.map((d) => d.field)}
              value={value.rows}
              onChange={(rows) => {
                // Mirror of the Columns handler above's own WHY comment — same exclusivity gap, same
                // fix, for the Rows side of the same "+" Add popover.
                const newlyAddedToRows = rows.filter((f) => !value.rows.includes(f));
                const columns = value.columns.filter((f) => !newlyAddedToRows.includes(f));
                // A field still present in Columns keeps its grouping.
                const stillGrouped = (field) => rows.includes(field) || columns.includes(field);
                const dateGroupings = value.dateGroupings.filter((g) => stillGrouped(g.field));
                const numericGroupings = value.numericGroupings.filter((g) => stillGrouped(g.field));
                const calculatedItems = clearOrphanedCalculatedItems(value.calculatedItems, rows, columns);
                // A Value's baseField (Show Values As PercentOf/DifferenceFrom/PercentDifferenceFrom)
                // targets a Rows field — removed here, same dangling-reference class as
                // clearDanglingMeasureReferences above, just for Rows instead of Values.
                const values = clearDanglingBaseFieldReferences(value.values, rows);
                // A field's own Field Settings Subtotals override is meaningless once it's left
                // Rows — same orphaned-per-field-config cleanup as dateGroupings/numericGroupings/
                // calculatedItems above, just for fieldSubtotals (added 2026-10-13).
                // (A Columns drilldown's own override stays — it isn't a Rows setting.)
                const fieldSubtotals = value.fieldSubtotals.filter((f) => rows.includes(f.field) || f.field === value.columnsHierarchyDrilldown?.topLevelField);
                // Same orphaned-per-field cleanup, for Show Items With No Data (PLAN 7, 2026-10-13).
                const fieldsShowingItemsWithNoData = value.fieldsShowingItemsWithNoData.filter((f) => rows.includes(f));
                // Drilldown v3: the drilldown keeps its slot among these Rows fields unless it was
                // removed or a same-hierarchy field joined (see reconcileDrilldowns — both axes).
                const next = {
                  ...value,
                  rows,
                  columns,
                  values,
                  dateGroupings,
                  numericGroupings,
                  calculatedItems,
                  fieldSubtotals,
                  fieldsShowingItemsWithNoData,
                  rowLevelSorts: [],
                  columnLevelSorts: newlyAddedToRows.length > 0 ? [] : value.columnLevelSorts,
                };
                onChange({ ...next, ...reconcileDrilldowns(next) });
              }}
              dragSourceKey="rows"
              onFieldDropped={moveField}
              // See the Columns instance's own WHY comment above — same reasoning, Rows also
              // accepts the Σ Values sentinel.
              isFieldAccepted={(field) => field === MEASURES_PLACEHOLDER || dimensionOptions.some((opt) => opt.value === field)}
              contextMenuTargets={[
                { zoneKey: "filters", label: "Move to Report Filter" },
                { zoneKey: "columns", label: "Move to Column Labels" },
                { zoneKey: "values", label: "Move to Values" },
              ]}
              // Real Excel's per-field "Field Settings…" (2026-10-13) — Subtotals only for now (see
              // PLAN 6's own scope cuts: per-field Layout needs more design work first, Subtotals'
              // "Custom" function picker is a separate follow-up). Rows-only: SubtotalsProcessor has
              // never applied Subtotals to Columns, so there's nothing for a Columns field's own
              // Field Settings to configure yet — that instance above gets no fieldSettings prop at
              // all, which hides the menu item entirely rather than showing a non-functional one.
              fieldSettings={{
                reportDefault: value.subtotals,
                overrides: value.fieldSubtotals,
                // patch is a partial { position, functions } — merged onto whatever override
                // already exists for this field (or defaults, for a field with none yet). An
                // override collapses back to "no entry at all" once both position and functions are
                // back to their Automatic/null state, same "don't store a no-op override" rule as
                // before this (PLAN 6 Phase 2) extended it from position-only to position+functions.
                onChange: (field, patch) => {
                  const existing = value.fieldSubtotals.find((f) => f.field === field);
                  const next = { position: existing?.position ?? null, functions: existing?.functions ?? null, ...patch };
                  const withoutField = value.fieldSubtotals.filter((f) => f.field !== field);
                  const isDefault = !next.position && (!next.functions || next.functions.length === 0);
                  onChange({
                    ...value,
                    fieldSubtotals: isDefault ? withoutField : [...withoutField, { field, ...next }],
                  });
                },
              }}
              // Real Excel's per-field "Field Settings…" > Custom Name / Layout & Print (PLAN 6
              // Phase 2, 2026-10-13) — a SEPARATE prop from fieldSettings above because this data
              // lives OUTSIDE `value`/`onChange` entirely (see PivotBuilder.jsx's own fieldLayouts
              // state WHY comment: none of these four change the query result, only the display, so
              // they're kept out of PivotRequest/structuralSignature on purpose).
              fieldLayout={{
                overrides: fieldLayouts,
                reportLayout,
                reportRepeatItemLabels: repeatItemLabels,
                onChange: (field, patch) => {
                  const existing = fieldLayouts.find((f) => f.field === field) ?? {};
                  const next = { ...existing, ...patch };
                  const withoutField = fieldLayouts.filter((f) => f.field !== field);
                  const isDefault = !next.customName && next.compact == null && next.repeatItemLabels == null && !next.insertBlankLine;
                  onFieldLayoutsChange(isDefault ? withoutField : [...withoutField, { field, ...next }]);
                },
              }}
              // Real Excel's per-field "Show items with no data" (PLAN 7, 2026-10-13, EXPERIMENTAL —
              // see the plan file's own WHY comment on why this needs live-cube verification). Back
              // in `value`/`onChange` (unlike fieldLayout above) — this one genuinely changes which
              // rows come back from the query, so it has to reach request.fieldsShowingItemsWithNoData,
              // not the presentation-only sibling list.
              fieldShowNoData={{
                fields: value.fieldsShowingItemsWithNoData,
                onChange: (field, checked) => {
                  const withoutField = value.fieldsShowingItemsWithNoData.filter((f) => f !== field);
                  onChange({
                    ...value,
                    fieldsShowingItemsWithNoData: checked ? [...withoutField, field] : withoutField,
                  });
                },
              }}
            />
          </div>

          {/* No separate aggregation dropdown: confirmed against the real cube (2026-09-21) that
              measures are pre-built with their aggregation baked in — see PivotValueField in the
              backend for the full story. Measures are individually draggable straight from the
              Available Fields panel, nested under their own measure group (2026-10-04, restructured
              2026-10-05 — see AvailableFieldsPanel's own WHY comment), same as any dimension field
              — the "+" Popover below is a faster way to add several at once. A pill's own Popover
              then covers Show Values As/Format/Conditional format, and pills can still be
              drag-reordered among themselves since that order affects the generated MDX (see
              handleValueRowDrop). */}
          <div className="flex flex-col rounded-xl border border-slate-200 bg-white p-2 shadow-sm">
            {/* Same header shape as Filters/Columns/Rows now — icon, label, count, a "+" Popover
                for the bulk-add path — rather than an always-visible "Choose measures…" control
                that made this box look and behave differently from the other three even though
                drag now works identically across all four (flagged live, 2026-10-04: "ليه جزء
                الـ measure دا مش زي الـ 3 التانيين؟"). */}
            <div className="mb-1.5 flex items-center justify-between gap-2">
              <span className="flex items-center gap-1.5 text-xs font-bold uppercase tracking-wide text-muted">
                <Sigma className="h-3.5 w-3.5" />
                Values {value.values.length > 0 && <span className="text-gold">({value.values.length})</span>}
              </span>
              <Popover
                align="right"
                trigger={(toggle) => (
                  <button
                    type="button"
                    onClick={toggle}
                    className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-gold/15 text-gold hover:bg-gold/25"
                    aria-label="Add a value"
                    title="Add a value"
                  >
                    <Plus className="h-3.5 w-3.5" />
                  </button>
                )}
              >
                {() => (
                  <div>
                    <p className="mb-2 text-xs text-muted">Choose one or more measures to add as Values.</p>
                    <Select
                      isMulti
                      className="react-select-container"
                      classNamePrefix="react-select"
                      placeholder="Choose measures…"
                      options={valueFieldOptions}
                      value={valueFieldOptions.filter((opt) => value.values.some((v) => v.field === opt.value))}
                      onChange={(selected) => setSelectedValueFields((selected ?? []).map((opt) => opt.value))}
                      menuPortalTarget={typeof document !== "undefined" ? document.body : null}
                      menuPosition="fixed"
                      styles={{ menuPortal: (base) => ({ ...base, zIndex: 10000 }) }}
                    />
                  </div>
                )}
              </Popover>
            </div>
            <ZoneDropArea
              zoneKey="values"
              onFieldDropped={moveField}
              // Real Excel's own behavior: a non-measure field dropped into Values still works —
              // it defaults to Count instead of Sum (confirmed live, 2026-10-05, against a normal
              // Excel/Power Pivot PivotTable). dimensionOptions covers this; the Σ Values sentinel
              // itself still isn't a real field, so it stays rejected (never in either list).
              isFieldAccepted={(field) => valueFieldOptions.some((opt) => opt.value === field) || dimensionOptions.some((opt) => opt.value === field)}
              isEmpty={value.values.length === 0}
            >
              <div className="flex flex-wrap gap-1.5">
                {value.values.map((v, index) => (
                  <ValuePill
                    key={index}
                    v={v}
                    rows={value.rows}
                    onUpdate={(patch, options) => updateValue(index, patch, options)}
                    onUpdateFormat={(patch) => updateValueFormat(index, patch)}
                    onRemove={() => removeValue(index)}
                    onDropReorder={(e) => handleValueRowDrop(e, index)}
                    onMoveTo={(zone) => moveField(v.field, "values", zone)}
                    onMoveUp={() => moveValueTo(index, index - 1)}
                    onMoveDown={() => moveValueTo(index, index + 1)}
                    onMoveToBeginning={() => moveValueTo(index, 0)}
                    onMoveToEnd={() => moveValueTo(index, value.values.length - 1)}
                    isFirst={index === 0}
                    isLast={index === value.values.length - 1}
                  />
                ))}
              </div>
            </ZoneDropArea>
          </div>
        </div>

        <AxisGroupingEditor
          label="Group Rows by"
          // Σ Values isn't a real field — "group it by Month/Quarter/Year" has no meaning, so it's
          // excluded here even when present in Rows (see AvailableFieldsPanel's own "Σ Values"
          // group, the field list entry you drag it in from).
          fields={value.rows.filter((f) => f !== MEASURES_PLACEHOLDER)}
          dimensions={dimensions}
          dateGroupings={value.dateGroupings}
          numericGroupings={value.numericGroupings}
          onChange={(dateGroupings, numericGroupings) => onChange({ ...value, dateGroupings, numericGroupings })}
        />
        <AxisGroupingEditor
          label="Group Columns by"
          fields={value.columns.filter((f) => f !== MEASURES_PLACEHOLDER)}
          dimensions={dimensions}
          dateGroupings={value.dateGroupings}
          numericGroupings={value.numericGroupings}
          onChange={(dateGroupings, numericGroupings) => onChange({ ...value, dateGroupings, numericGroupings })}
        />
      </div>
    </div>
  );
}
