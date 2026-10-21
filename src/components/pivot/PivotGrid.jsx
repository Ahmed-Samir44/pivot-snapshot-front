import { Fragment, useEffect, useImperativeHandle, useMemo, useRef, useState } from "react";
import { ChevronDown, ChevronRight, ChevronLeft, FileSpreadsheet, Filter as FilterIcon, Loader2, Table2, X } from "lucide-react";
import Popover from "../forms/Popover";
import ContextMenu from "../forms/ContextMenu";
import Modal from "../forms/Modal";
import Slicer from "./Slicer";
import Timeline from "./Timeline";
import GroupItemsPanel from "./GroupItemsPanel";
import { FieldSortDialog, FilterRow, NumberFormatEditor, ValueFieldSettingsPanel, ConditionalFormatEditor, SHOW_VALUES_AS_OPTIONS, DEFAULT_FORMAT } from "./FieldPicker";
import { formatForShowValuesAs } from "../../utils/showValuesAsFormat";
import { formatForColumn } from "../../utils/numberFormat";
import { DEFAULT_PIVOT_OPTIONS, cellText, filterAreaSlot, filterValueSummary } from "../../utils/pivotOptions";
import { filterKind, setFieldFilter, FILTER_KIND_LABELS } from "../../utils/filterKinds";
import { LABEL_FILTERS, VALUE_FILTERS, DATE_FILTERS, isLabelMode, isDateMode, datePeriodLabel, isRangeMode, isTop10Mode, top10Parts, isFilterComplete } from "../../utils/filterModes";
import { formatDimensionName } from "../../utils/cubeMeta";
import { conditionalStyleFor, detailRangeForColumn, detailRangeForRow, iconColorFor } from "../../utils/conditionalFormat";
import { buildRowHeaderGroups, buildDisplayRows, applyCollapse, collapsibleKeys, collapseKey, buildColumnHeaderRows, computeLabelSpans } from "../../utils/pivotGridLayout";
import { MEASURES_PLACEHOLDER, measuresIndexOf, buildCaptionToIndex, measureIndexOf } from "../../utils/measuresPlacement";
import { applyReportSubtotals } from "../../utils/subtotals";
import { drillthroughPivot } from "../../services/pivotApi";
import { removeFieldEverywhere } from "../../utils/dragDrop";
import { selectionStats } from "../../utils/selectionStats";
import { cellFormatKey } from "../../utils/cellFormats";
import { sharedFormatType } from "../../utils/numberFormatCommands";

// The drag payload type for moving a row item inside the grid (Excel's drag-an-item-to-reorder).
const GRID_ITEM_DRAG_MIME = "application/x-pivot-grid-item";

// A one-line description of a filter for the field menu's list of active filters, from the same
// mode names the filter editor uses (utils/filterModes.js).
function describeFilter(filter) {
  const mode = filter.mode ?? "Members";
  if (mode === "Members") return filter.includedMembers.length === 1 ? filter.includedMembers[0] : `${filter.includedMembers.length} items`;
  if (mode === "Exclude") return `all but ${filter.includedMembers.length}`;
  const name = (list) => list.find((f) => f.value === mode)?.label.replace("…", "").toLowerCase() ?? mode;
  if (isLabelMode(mode)) return `${name(LABEL_FILTERS)} "${filter.labelText}"${isRangeMode(mode) ? ` and "${filter.labelTextTo}"` : ""}`;
  if (mode === "DatePeriod") return datePeriodLabel(filter.labelText).toLowerCase();
  if (isDateMode(mode)) return `${name(DATE_FILTERS)} ${filter.labelText}${isRangeMode(mode) ? ` and ${filter.labelTextTo}` : ""}`;
  if (isTop10Mode(mode)) {
    const { direction, unit } = top10Parts(mode);
    return `${direction} ${unit === "Items" ? filter.n : filter.value}${unit === "Percent" ? "%" : unit === "Sum" ? " (sum)" : ""} by ${filter.byMeasureField}`;
  }
  return `${filter.byMeasureField} ${name(VALUE_FILTERS)} ${filter.value}${isRangeMode(mode) ? ` and ${filter.valueTo}` : ""}`;
}

// Excel's own Report Filter area: every field in Filters gets its own row ABOVE the PivotTable
// itself — field name on the left, current selection (plus a dropdown to change it) on the right
// — confirmed live 2026-10-05 against a real Excel screenshot. This was previously only editable
// from the separate Filters box in the field-picker sidebar; Excel shows (and lets you re-filter
// from) it right here too, which this app didn't have at all.
//
// The selection caption ("(All)", the member, "(Multiple Items)") is pivotOptions.js's
// filterValueSummary, shared with the export and mirrored by the saved snapshot. `style` places the
// row in the filter area's grid (PivotTable Options > Layout & Format, filterAreaSlot).
function FilterHeaderRow({ filter, index, request, onChange, dimensionOptions, selectedValueOptions, style }) {
  const updateFilter = (patch) => {
    const next = [...request.filters];
    next[index] = { ...next[index], ...patch };
    onChange({ ...request, filters: next });
  };
  const removeFilter = () => {
    onChange({ ...request, filters: request.filters.filter((_, i) => i !== index) });
  };

  return (
    <div className="flex items-stretch overflow-hidden rounded-lg border border-slate-200 bg-white" style={style}>
      <div className="flex w-40 shrink-0 items-center border-r border-slate-200 bg-slate-50 px-3 py-1.5 text-xs font-semibold text-ink">
        <span className="truncate">{formatDimensionName(filter.field)}</span>
      </div>
      <Popover
        fullWidth
        trigger={(toggle) => (
          <button
            type="button"
            onClick={toggle}
            className="flex min-w-[10rem] flex-1 items-center justify-between gap-2 px-3 py-1.5 text-left text-xs text-ink hover:bg-slate-50"
          >
            <span className="truncate">{filterValueSummary(filter)}</span>
            <FilterIcon className="h-3 w-3 shrink-0 text-gold" />
          </button>
        )}
      >
        {(close) => (
          <div className="space-y-2">
            {/* Explicit close button, matching FieldPicker.jsx's own FilterPill popover — requested
                live, 2026-10-06 ("دي ليه مافيش زرار بيقفلها"), same reasoning there applies here. */}
            <div className="flex justify-end">
              <button type="button" onClick={close} className="shrink-0 rounded-full p-1 text-muted hover:bg-slate-100 hover:text-ink" aria-label="Close">
                <X className="h-4 w-4" />
              </button>
            </div>
            <FilterRow
              filter={filter}
              dimensionOptions={dimensionOptions}
              selectedValueOptions={selectedValueOptions}
              onChange={updateFilter}
              onRemove={removeFilter}
              onClose={close}
              hideFieldPicker={Boolean(filter.field)}
            />
          </div>
        )}
      </Popover>
    </div>
  );
}

// Flat, bordered Excel-style header cell — replaces the earlier rounded "pill" chrome (explicitly
// NOT a traditional spreadsheet header cell, per that version's own comment) with something that
// actually reads as a PivotTable field header: tinted background, bold text, real borders on every
// edge (see the <th>/<td> className strings below for the gridlines themselves), requested live
// 2026-09-28 ("نفس الشكل نفس الـ behaviours" — look AND behave like Excel's own PivotTable).
// alignLeft: an outer column header spanning its children ("2024" over its months) is left-aligned,
// like Excel's Compact layout (2026-10-15, the user's call) — single-column cells stay centered.
// Reports the selected range's stats up to the status bar when they change. A child component so its
// effect can live past PivotGrid's no-result early return (no hook may sit after one).
// Exposes this render's commands on a ref the parent owns (React's useImperativeHandle) — a child
// so the hook can live past PivotGrid's no-result early return.
function CommandBinder({ target, value }) {
  useImperativeHandle(target, () => value);
  return null;
}

function SelectionStatsReporter({ stats, onChange }) {
  const key = JSON.stringify(stats);
  useEffect(() => {
    onChange(key === "null" ? null : JSON.parse(key));
  }, [key, onChange]);
  return null;
}

function HeaderCell({ children, muted = false, isTotal = false, alignLeft = false }) {
  // Neutral slate header band (see headerBg) — ink text; muted (field-name labels, not real member
  // captions) is a lighter slate, and a Total column header is bold ink with a gold underline.
  return (
    <span
      className={`flex items-center ${alignLeft ? "justify-start" : "justify-center"} gap-1 text-sm ${
        muted ? "font-medium text-slate-500" : isTotal ? "font-bold text-ink underline decoration-gold decoration-2 underline-offset-4" : "font-semibold text-ink"
      }`}
    >
      {children}
    </span>
  );
}

// A real, working field-header dropdown (Excel's own filter/sort arrow on each PivotTable field
// button) — not the old decorative, non-functional ChevronDown. `field` is the raw identifier
// (e.g. "[Dim Doctor].[Specialty Name]"), null when this header doesn't correspond to exactly one
// field (Compact's merged multi-field row header, or a multi-field Columns breadcrumb) — those
// render as a plain, non-interactive HeaderCell instead, see callers below.
//
// Sort here is AXIS-wide (same single-axis Sort the toolbar's own Sort popover sets in
// FieldPicker.jsx), not per-level — clicking any row level's arrow sorts the whole Rows axis, same
// as clicking it from any other level. Genuinely useful (instant access to the single most common
// action) without duplicating the separate per-level Sort UI that already exists in FieldPicker's
// toolbar for the 2+-fields-per-axis case. This mutates the LIVE `request` (same as every other
// Sort/Filter control) — it still needs Run query to take effect, same as the rest of this app.
// `level` is this field's 0-based position within its own axis (Columns only ever renders this
// menu at all when it has exactly one field, i.e. level 0 — see the Columns call site below; Rows
// passes its own real rowLevel). Needed so "Clear sort"/its own sorted-indicator can see a
// PivotBuilder-defaulted (or LevelSortEditor-set) per-level Ascending rule, not just the single
// shared `sort` field — see setSort's own WHY comment for why both mechanisms can be active.
// sortOnly (Drilldown v2, 2026-10-14): a drilldown hierarchy's "Row Labels" header offers just Sort
// A→Z / Z→A (sorted within each parent, see MdxPivotQueryBuilder.BuildHierarchyDrilldownSet). No
// Filter — filtering the same hierarchy that's on the axis would put it in WHERE too, which SSAS
// rejects — and no More Sort Options (its per-level editor needs real Rows fields); sorting by a
// value works from a value cell's right-click menu.
// fieldChoices ([{ field, label, level }], 2026-10-18): Excel's Compact-form "Row Labels" header —
// one arrow for several fields, with a "Select field:" box on top of its menu choosing which field
// every item below applies to (found missing by the full-app test). Replaces `field`/`level`.
function FieldHeaderMenu({ label, field: fixedField, axis, request, onChange, muted = false, isTotal = false, level: fixedLevel = 0, sortOnly = false, fieldChoices = null, fieldCollapse = null }) {
  // "More Sort Options…"/"Field Settings…" open the SAME panel in place of the plain shortcut
  // list, same back-and-forth pattern as Excel's own header dropdown (its "More Sort Options..."
  // item opens a bigger dialog from the same arrow) — relocated here from FieldPicker's toolbar
  // entirely, since Excel never puts per-field Sort/"Show items with no data" in a side task pane,
  // only on the field's own dropdown arrow or its right-click menu (moved 2026-10-04, "المفروض
  // تبقى في نفس مكان الـ excel").
  const [panel, setPanel] = useState("shortcuts");
  const [choiceIndex, setChoiceIndex] = useState(0);
  const choice = fieldChoices?.length ? fieldChoices[Math.min(choiceIndex, fieldChoices.length - 1)] : null;
  const field = choice ? choice.field : fixedField;
  const level = choice ? choice.level : fixedLevel;

  // Σ Values has no members/label of its own to sort or filter by — same restriction the backend
  // enforces for LevelSort (see MdxPivotQueryBuilder.ValidateAxisLevelSorts) — so it renders as a
  // plain, non-interactive header, same as the "no single field" case below.
  if (!field || field === MEASURES_PLACEHOLDER || !request || !onChange) {
    return <HeaderCell muted={muted} isTotal={isTotal}>{label}</HeaderCell>;
  }

  const levelSorts = axis === "Columns" ? request.columnLevelSorts : request.rowLevelSorts;
  const levelSort = levelSorts.find((s) => s.level === level);
  const isLevelSortedHere = Boolean(levelSort) && levelSort.byMeasureField == null;
  const isSortedHere = (request.sort?.axis === axis && !request.sort?.byMeasureField && !request.sort?.customOrder) || isLevelSortedHere;
  const hasFilter = request.filters.some((f) => f.field === field);
  const hasGroups = (request.itemGroups ?? []).some((g) => g.field === field);
  const noDataKey = axis === "Columns" ? "Columns" : "Rows";
  const showsNoDataItems = request.showItemsWithNoData.includes(noDataKey);

  const setSort = (direction) => {
    // `sort` and the per-field LevelSorts only conflict on the SAME axis (MdxPivotQueryBuilder.
    // ValidateLevelSorts, since 2026-10-15 — it used to be global, so a sort on one axis wiped the
    // other axis's rules, a drilldown's sort included). Each path below clears this axis only.
    //
    // Excel sorts the FIELD you clicked (2026-10-14, its per-field Sort dialog): a field on a
    // multi-field axis gets its own LevelSort — other fields' rules are kept — instead of the
    // axis-wide sort, which only works with one field on the axis. A drilldown slot (sortOnly) keeps
    // the axis-wide sort: the hierarchy is the axis's one sortable field there.
    if (sortOnly) {
      onChange({ ...request, sort: { direction, byMeasureField: null, axis, customOrder: null }, [axis === "Columns" ? "columnLevelSorts" : "rowLevelSorts"]: [] });
      return;
    }
    const key = axis === "Columns" ? "columnLevelSorts" : "rowLevelSorts";
    onChange({ ...request, sort: request.sort?.axis === axis ? null : request.sort, [key]: [...levelSorts.filter((s) => s.level !== level), { level, direction, byMeasureField: null }] });
  };

  const clearSort = () => {
    if (isLevelSortedHere) {
      const key = axis === "Columns" ? "columnLevelSorts" : "rowLevelSorts";
      onChange({ ...request, [key]: levelSorts.filter((s) => s.level !== level) });
      return;
    }
    onChange({ ...request, sort: null });
  };

  // "Allow multiple filters per field" (PivotTable Options > Totals & Filters, 2026-10-19): the field
  // can hold one filter of each kind (utils/filterKinds.js); off, one per field.
  const allowMultipleFilters = Boolean(request.allowMultipleFiltersPerField);
  const fieldFilters = request.filters.filter((f) => f.field === field);
  // The filter the panel opens on: the manual selection if there is one, else the field's first.
  const existingFilter = fieldFilters.find((f) => filterKind(f) === 0) ?? fieldFilters[0] ?? null;

  // Reused by FilterRow's own OK (fieldLocked=true, see its own WHY comment) — adds a new filter
  // entry for this field, or replaces the existing one: the field's only filter, or with multiple
  // filters allowed, its filter of the same kind (setFieldFilter).
  const applyFieldFilter = (updatedFilter) => {
    onChange({ ...request, filters: setFieldFilter(request.filters, updatedFilter, allowMultipleFilters) });
  };

  // "Clear Filter" clears every filter on the field, like Excel's "Clear Filter From <field>".
  const removeFilter = () => onChange({ ...request, filters: request.filters.filter((f) => f.field !== field) });
  // FilterRow's own remove: with multiple filters allowed, just the filter of that kind.
  const removeFilterOfKind = (kind) =>
    onChange({ ...request, filters: request.filters.filter((f) => f.field !== field || (allowMultipleFilters && filterKind(f) !== kind)) });

  // TopN/BottomN/GreaterThan/etc.'s own "Choose a measure…" dropdown needs this same shape
  // FieldPicker.jsx's Filters zone already builds from value.values — mirrored here since
  // FieldHeaderMenu only ever receives `request`, not a separate measureOptions prop.
  const selectedValueOptions = request.values.filter((v) => v.field).map((v) => ({ value: v.field, label: v.field }));

  const toggleNoDataItems = (checked) =>
    onChange({
      ...request,
      showItemsWithNoData: checked
        ? [...request.showItemsWithNoData, noDataKey]
        : request.showItemsWithNoData.filter((a) => a !== noDataKey),
    });

  return (
    <Popover
      align="right"
      fullWidth
      // The Sort dialog needs more room than the shortcut list.
      panelClassName={panel === "sort" || panel === "group" ? "min-w-[20rem] max-w-[22rem]" : "min-w-[14rem] max-w-[18rem]"}
      trigger={(toggle) => (
        <button
          type="button"
          onClick={() => {
            setPanel("shortcuts");
            toggle();
          }}
          // Same tones as HeaderCell on the neutral slate band (the 2026-10-13/14 contrast bugs came
          // from the old gold band: gray and gold text both vanished on it). Gold is the hover/
          // filtered accent only.
          className={`flex w-full items-center justify-center gap-1 text-sm font-semibold hover:text-gold-strong ${
            muted ? "text-slate-500" : isTotal ? "text-gold-strong" : "text-ink"
          }`}
          title={`Sort or filter ${label}`}
        >
          {label}
          {hasFilter && <FilterIcon className="h-3 w-3 shrink-0 text-gold-strong" />}
          <ChevronDown className="h-3.5 w-3.5 shrink-0" />
        </button>
      )}
    >
      {(close) => (
        <>
        {fieldChoices?.length > 1 && (
          <label className="mb-2 flex items-center gap-2 border-b border-slate-100 pb-2 text-xs font-medium text-ink">
            Select field:
            <select
              className="flex-1 rounded-lg border border-slate-200 bg-white px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-gold"
              value={Math.min(choiceIndex, fieldChoices.length - 1)}
              onChange={(e) => {
                setChoiceIndex(Number(e.target.value));
                setPanel("shortcuts");
              }}
            >
              {fieldChoices.map((c, i) => (
                <option key={c.field} value={i}>
                  {c.label}
                </option>
              ))}
            </select>
          </label>
        )}
        {panel === "sort" ? (
          <div>
            {/* Excel's per-field "Sort (<field>)" dialog (see FieldSortDialog). */}
            <FieldSortDialog key={field} field={field} axis={axis} level={level} request={request} onChange={onChange} onBack={() => setPanel("shortcuts")} onDone={close} />
          </div>
        ) : panel === "group" ? (
          <GroupItemsPanel key={field} field={field} request={request} onChange={onChange} onBack={() => setPanel("shortcuts")} onDone={close} />
        ) : panel === "filter" ? (
          <div>
            <button
              type="button"
              onClick={() => setPanel("shortcuts")}
              className="mb-2 inline-flex items-center gap-1 text-xs font-medium text-muted hover:text-ink"
            >
              <ChevronLeft className="h-3.5 w-3.5" /> Back
            </button>
            {/* This field is already fixed (whichever header's arrow was clicked) — no field
                picker, and Label/Value Filters ("More Filters…")/Exclude become available, matching
                real Excel's own Row/Column field dropdown (confirmed live 2026-10-12) — see
                FilterRow's own WHY comment for the full Report-Filter-vs-field-dropdown story. */}
            {/* With multiple filters per field, the field's active filters (one per kind), each
                removable on its own; the editor below adds or replaces the kind it's set to. */}
            {allowMultipleFilters && fieldFilters.length > 0 && (
              <ul className="mb-2 space-y-1 border-b border-slate-100 pb-2">
                {fieldFilters.map((f) => (
                  <li key={filterKind(f)} className="flex items-center justify-between gap-2 text-xs text-ink">
                    <span>
                      <span className="font-medium">{FILTER_KIND_LABELS[filterKind(f)]}:</span> {describeFilter(f)}
                    </span>
                    <button
                      type="button"
                      onClick={() => removeFilterOfKind(filterKind(f))}
                      className="rounded-full p-0.5 text-muted hover:bg-slate-100 hover:text-ink"
                      aria-label={`Remove ${FILTER_KIND_LABELS[filterKind(f)]}`}
                    >
                      <X className="h-3 w-3" />
                    </button>
                  </li>
                ))}
              </ul>
            )}
            <FilterRow
              key={field}
              filter={existingFilter ?? { field, mode: "Members", includedMembers: [] }}
              dimensionOptions={[]}
              selectedValueOptions={selectedValueOptions}
              onChange={applyFieldFilter}
              onRemove={() => (allowMultipleFilters ? removeFilterOfKind(filterKind(existingFilter ?? { mode: "Members" })) : removeFilter())}
              onClose={close}
              hideFieldPicker
              showAdvancedFilters
            />
          </div>
        ) : (
          <div className="flex flex-col gap-1 text-sm">
            <button
              type="button"
              onClick={() => {
                setSort("Ascending");
                close();
              }}
              className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10"
            >
              Sort A → Z
            </button>
            <button
              type="button"
              onClick={() => {
                setSort("Descending");
                close();
              }}
              className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10"
            >
              Sort Z → A
            </button>
            {isSortedHere && (
              <button
                type="button"
                onClick={() => {
                  clearSort();
                  close();
                }}
                className="rounded-lg px-2 py-1.5 text-left text-muted hover:bg-slate-100"
              >
                Clear sort
              </button>
            )}
            {!sortOnly && (
              <>
            <button type="button" onClick={() => setPanel("sort")} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
              More Sort Options…
            </button>
            <div className="my-1 border-t border-slate-100" />
            {hasFilter && (
              <button
                type="button"
                onClick={() => {
                  removeFilter();
                  close();
                }}
                className="rounded-lg px-2 py-1.5 text-left text-muted hover:bg-slate-100"
              >
                Clear Filter From &quot;{choice?.label ?? label}&quot;
              </button>
            )}
            <button type="button" onClick={() => setPanel("filter")} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
              Filter…
            </button>
            {/* Excel's Expand/Collapse Entire Field (2026-10-20) — only for a field whose items have
                a subtotal row to collapse into (the outer Rows fields). */}
            {fieldCollapse?.canCollapse(level) && (
              <>
                <div className="my-1 border-t border-slate-100" />
                <button
                  type="button"
                  onClick={() => {
                    fieldCollapse.expand(level);
                    close();
                  }}
                  disabled={!fieldCollapse.isCollapsed(level)}
                  className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10 disabled:cursor-not-allowed disabled:opacity-40"
                >
                  Expand Entire Field
                </button>
                <button
                  type="button"
                  onClick={() => {
                    fieldCollapse.collapse(level);
                    close();
                  }}
                  className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10"
                >
                  Collapse Entire Field
                </button>
                <div className="my-1 border-t border-slate-100" />
              </>
            )}
            {/* Excel's Group Selection / Ungroup (2026-10-20) — see GroupItemsPanel. */}
            <button type="button" onClick={() => setPanel("group")} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
              {hasGroups ? "Groups…" : "Group Items…"}
            </button>
            <div className="my-1 border-t border-slate-100" />
            <label className="flex items-center gap-2 px-2 py-1.5 font-medium text-ink">
              <input
                type="checkbox"
                className="h-4 w-4 accent-gold"
                checked={showsNoDataItems}
                onChange={(e) => toggleNoDataItems(e.target.checked)}
              />
              Show items with no data
            </label>
              </>
            )}
          </div>
        )}
        </>
      )}
    </Popover>
  );
}

// Excel's own right-click menu on a PivotTable VALUE cell (Sort/Number Format/Show Values As/Remove
// field/Show Details) — confirmed live 2026-10-05 against a real Excel screenshot naming these as
// the ones "people actually use" (اعمله زي الإكسل بالظبط). Sort/Number Format/Show Values As/Remove
// reuse the exact same request.values[index] patch shapes FieldPicker.jsx's ValuePill already
// applies from the Values sidebar — this is just a second entry point onto the same live request,
// not new business logic. Show Details is genuinely new (MDX DRILLTHROUGH via
// services/pivotApi.js's drillthroughPivot) — only offered when this cell has real member identity
// on both axes (row.uniqueNames/column.uniqueNames — null for a Grand Total/Subtotal row) and the
// value isn't a calculated field, a dimension-aggregation value, or ShowValuesAs-transformed (none
// of those are a single stored fact-level quantity DRILLTHROUGH can meaningfully target — see
// MdxPivotQueryBuilder.BuildDrillthrough's own WHY comment on the backend).
// A field-name header cell with field captions turned off (PivotTable Options > Display): empty.
function NoCaption() {
  return null;
}

// Whether a value cell has underlying rows Show Details can fetch: real members on both axes (a
// total has none), a plain measure (no Show Values As, not a calculated field or a "Count of
// <field>"). A Calculated Item is a synthetic MDX member (e.g. "East + West" combined into one
// row/column) with a real, non-null UniqueName of its own — NOT backed by a single real fact-table
// row set the same way a Calculated Field isn't, so it needs the exact same exclusion (caught
// 2026-10-06, full codebase review: Show Details was still offered on a Calculated Item cell, whose
// drillthrough would at best return only one of the combined members' own rows). Shared by the
// right-click menu and double-click.
function cellHasDetails(request, valueField, row, column) {
  const isDimensionField = valueField.field.startsWith("[");
  const isCalculatedField = request.calculatedFields.some((cf) => cf.name === valueField.field);
  const isCalculatedItem = Boolean(row.isCalculatedItem) || Boolean(column.isCalculatedItem);
  // A Group Selection item is a calculated member too — SSAS can't drill through it (2026-10-20).
  const groupNames = (request.itemGroups ?? []).map((g) => `.[${g.name}]`);
  const isGroupItem = [...(row.uniqueNames ?? []), ...(column.uniqueNames ?? [])].some((u) => groupNames.some((n) => u?.endsWith(n)));
  // A value saved without showValuesAs is a plain one (the backend's default is Normal too).
  return (
    Boolean(row.uniqueNames) &&
    Boolean(column.uniqueNames) &&
    (valueField.showValuesAs ?? "Normal") === "Normal" &&
    !isDimensionField &&
    !isCalculatedField &&
    !isCalculatedItem &&
    !isGroupItem
  );
}

// Excel's right-click menu on a ROW LABEL (2026-10-20): Copy, Sort, Filter (Keep Only / Hide /
// Clear Filter From), Expand/Collapse (the item, the entire field), Move (Up / Down / To Beginning /
// To End — a manual order, LevelSort.CustomOrder), Remove the field. Flat sections instead of
// Excel's flyout submenus. Refresh, PivotTable Options… and Hide Field List too (2026-10-21), as on
// Excel's own row-label menu — the value-cell menu already had them.
function MenuSection({ children }) {
  return <p className="px-2 pb-0.5 pt-1.5 text-[11px] font-semibold uppercase tracking-wide text-muted">{children}</p>;
}

function RowLabelContextMenu({ x, y, level, label, filterable = true, field, fieldLabel, request, onChange, onClose, fieldCollapse, onToggleItem, itemCollapsed, onMove, onRefresh = null, onOpenPivotTableOptions = null, onHideFieldList = null }) {
  const allowMultipleFilters = Boolean(request.allowMultipleFiltersPerField);
  const manual = request.filters.find((f) => f.field === field && filterKind(f) === 0);
  const hasFilter = request.filters.some((f) => f.field === field);
  const run = (action) => () => {
    action();
    onClose();
  };
  const setSort = (direction) =>
    onChange({
      ...request,
      sort: request.sort?.axis === "Rows" ? null : request.sort,
      rowLevelSorts: [...request.rowLevelSorts.filter((s) => s.level !== level), { level, direction, byMeasureField: null }],
    });
  const keepOnly = () => onChange({ ...request, filters: setFieldFilter(request.filters, { field, mode: "Members", includedMembers: [label] }, allowMultipleFilters) });
  // Hide: out of the manual selection if there is one, else excluded.
  const hide = () => {
    const next =
      manual?.mode === "Members" && manual.includedMembers.length > 0
        ? { ...manual, includedMembers: manual.includedMembers.filter((m) => m !== label) }
        : { field, mode: "Exclude", includedMembers: [...new Set([...(manual?.mode === "Exclude" ? manual.includedMembers : []), label])] };
    onChange({ ...request, filters: setFieldFilter(request.filters, next, allowMultipleFilters) });
  };
  const itemClass = "rounded-lg px-2 py-1.5 text-left hover:bg-gold/10 disabled:cursor-not-allowed disabled:opacity-40";
  const short = label.length > 24 ? `${label.slice(0, 23)}…` : label;

  return (
    <ContextMenu x={x} y={y} onClose={onClose}>
      <div className="flex flex-col">
        <button type="button" className={itemClass} onClick={run(() => navigator.clipboard?.writeText(label).catch(() => {}))}>
          Copy
        </button>
        {onRefresh && (
          <button type="button" className={itemClass} onClick={run(onRefresh)}>
            Refresh
          </button>
        )}
        <MenuSection>Sort</MenuSection>
        <button type="button" className={itemClass} onClick={run(() => setSort("Ascending"))}>
          Sort A → Z
        </button>
        <button type="button" className={itemClass} onClick={run(() => setSort("Descending"))}>
          Sort Z → A
        </button>
        <MenuSection>Filter</MenuSection>
        <button type="button" className={itemClass} disabled={!filterable} onClick={run(keepOnly)}>
          Keep Only Selected Items
        </button>
        {/* Hiding the only item left would empty the selection — which means "no filter", so every
            item would come back. Excel greys it out there too. */}
        <button
          type="button"
          className={itemClass}
          disabled={!filterable || (manual?.mode === "Members" && manual.includedMembers.length === 1 && manual.includedMembers[0] === label)}
          onClick={run(hide)}
        >
          Hide Selected Items
        </button>
        <button type="button" className={itemClass} disabled={!hasFilter} onClick={run(() => onChange({ ...request, filters: request.filters.filter((f) => f.field !== field) }))}>
          Clear Filter From &quot;{fieldLabel}&quot;
        </button>
        {(onToggleItem || fieldCollapse?.canCollapse(level)) && (
          <>
            <MenuSection>Expand/Collapse</MenuSection>
            {onToggleItem && (
              <button type="button" className={itemClass} onClick={run(onToggleItem)}>
                {itemCollapsed ? "Expand" : "Collapse"}
              </button>
            )}
            {fieldCollapse?.canCollapse(level) && (
              <>
                <button type="button" className={itemClass} disabled={!fieldCollapse.isCollapsed(level)} onClick={run(() => fieldCollapse.expand(level))}>
                  Expand Entire Field
                </button>
                <button type="button" className={itemClass} onClick={run(() => fieldCollapse.collapse(level))}>
                  Collapse Entire Field
                </button>
              </>
            )}
          </>
        )}
        <MenuSection>Move</MenuSection>
        <button type="button" className={itemClass} onClick={run(() => onMove("beginning"))}>
          Move &quot;{short}&quot; to Beginning
        </button>
        <button type="button" className={itemClass} onClick={run(() => onMove("up"))}>
          Move &quot;{short}&quot; Up
        </button>
        <button type="button" className={itemClass} onClick={run(() => onMove("down"))}>
          Move &quot;{short}&quot; Down
        </button>
        <button type="button" className={itemClass} onClick={run(() => onMove("end"))}>
          Move &quot;{short}&quot; to End
        </button>
        <div className="my-1 border-t border-slate-100" />
        <button type="button" className={itemClass} onClick={run(() => onChange(removeFieldEverywhere(request, field)))}>
          Remove &quot;{fieldLabel}&quot;
        </button>
        {(onOpenPivotTableOptions || onHideFieldList) && <div className="my-1 border-t border-slate-100" />}
        {onOpenPivotTableOptions && (
          <button type="button" className={itemClass} onClick={run(onOpenPivotTableOptions)}>
            PivotTable Options…
          </button>
        )}
        {onHideFieldList && (
          <button type="button" className={itemClass} onClick={run(onHideFieldList)}>
            Hide Field List
          </button>
        )}
      </div>
    </ContextMenu>
  );
}

// showDetailsEnabled: PivotTable Options > Data > "Enable show details" — off, Show Details stays
// in the menu but grayed out, like Excel.
function ValueCellContextMenu({ x, y, onClose, request, onChange, valueField, measureIndex, sortAxis, row, column, onShowDetails, onRefresh, displayValue, onHideFieldList, onOpenPivotTableOptions, showDetailsEnabled = true, onCopyRange = null }) {
  const [panel, setPanel] = useState("root");
  const format = valueField.format ?? DEFAULT_FORMAT;
  const isDimensionField = valueField.field.startsWith("[");
  const displayLabel = valueField.customLabel || (isDimensionField ? formatDimensionName(valueField.field) : valueField.field);

  const updateValue = (patch) => {
    const values = request.values.map((v, i) => (i === measureIndex ? { ...v, ...patch } : v));
    onChange({ ...request, values });
  };
  const updateFormat = (patch) => updateValue({ format: { ...format, ...patch } });
  // Same global-exclusivity clear as FieldHeaderMenu's own setSort above — this menu's Sort
  // shortcut hits the identical "both axes default-sorted, this throws otherwise" case.
  const setSort = (direction) =>
    onChange({ ...request, sort: { direction, byMeasureField: valueField.field, axis: sortAxis, customOrder: null }, [sortAxis === "Columns" ? "columnLevelSorts" : "rowLevelSorts"]: [] });
  const removeValue = () => onChange({ ...request, values: request.values.filter((_, i) => i !== measureIndex) });

  const canShowDetails = cellHasDetails(request, valueField, row, column);

  return (
    <ContextMenu x={x} y={y} onClose={onClose}>
      {panel === "root" ? (
        <div className="flex flex-col gap-1">
          <button
            type="button"
            onClick={async () => {
              onClose();
              // A multi-cell selection is copied whole (tab-separated, like Excel).
              if (onCopyRange) {
                onCopyRange();
                return;
              }
              try {
                await navigator.clipboard.writeText(displayValue);
              } catch {
                // Clipboard access can be denied (permissions, insecure context) — nothing useful
                // to surface to the user for a plain copy-to-clipboard convenience action.
              }
            }}
            className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10"
          >
            Copy
          </button>
          {/* Disabled, not hidden — matches Excel's own menu item-for-item, but this report has no
              per-cell font/border/fill styling the way a generic Excel sheet does: NumberFormat
              (below) and Conditional Formatting (also below) are the two pieces of "Format Cells"
              that mean anything for an aggregated, re-queried pivot — a manual per-cell style would
              just get wiped out the next time this cell's query reruns anyway. */}
          <button
            type="button"
            disabled
            title="Not available — use Number Format and Conditional Formatting below instead (this report has no separate per-cell styling)."
            className="cursor-not-allowed rounded-lg px-2 py-1.5 text-left text-muted/50"
          >
            Format Cells…
          </button>
          <button type="button" onClick={() => setPanel("numberFormat")} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
            Number Format…
          </button>
          <button type="button" onClick={() => { onClose(); onRefresh(); }} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
            Refresh
          </button>
          <button
            type="button"
            onClick={() => setPanel("sort")}
            className="flex items-center justify-between gap-2 rounded-lg px-2 py-1.5 text-left hover:bg-gold/10"
          >
            Sort <ChevronRight className="h-3.5 w-3.5 shrink-0" />
          </button>
          {/* Quick Explore drills into RELATED tables in Excel's own Data Model/Power Pivot — this
              report queries one SSAS cube directly, with no separate relationship graph to explore
              the same way. */}
          <button
            type="button"
            disabled
            title="Not available — this report queries one cube directly, with no separate related-table graph to drill into."
            className="cursor-not-allowed rounded-lg px-2 py-1.5 text-left text-muted/50"
          >
            Quick Explore
          </button>
          <button type="button" onClick={() => { removeValue(); onClose(); }} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
            Remove &quot;{displayLabel}&quot;
          </button>
          <button
            type="button"
            onClick={() => setPanel("showValuesAs")}
            className="flex items-center justify-between gap-2 rounded-lg px-2 py-1.5 text-left hover:bg-gold/10"
          >
            Show Values As <ChevronRight className="h-3.5 w-3.5 shrink-0" />
          </button>
          {canShowDetails && (
            <button
              type="button"
              disabled={!showDetailsEnabled}
              title={showDetailsEnabled ? undefined : "Turned off in PivotTable Options > Data > Enable show details."}
              onClick={() => { onClose(); onShowDetails(); }}
              className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10 disabled:cursor-not-allowed disabled:text-muted/50 disabled:hover:bg-transparent"
            >
              Show Details
            </button>
          )}
          {/* Excel's own "Additional Actions" here is populated by Power Pivot/Data Model features
              (Edit Data Model, What-If Analysis) — none of which exist for a live SSAS cube report. */}
          <button
            type="button"
            disabled
            title="Not available — Excel's own Additional Actions here are Power Pivot/Data Model features, which don't apply to a live cube report."
            className="flex cursor-not-allowed items-center justify-between gap-2 rounded-lg px-2 py-1.5 text-left text-muted/50"
          >
            Additional Actions <ChevronRight className="h-3.5 w-3.5 shrink-0" />
          </button>
          <button type="button" onClick={() => setPanel("settings")} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
            Value Field Settings…
          </button>
          <button type="button" onClick={() => setPanel("conditionalFormat")} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
            Conditional Formatting…
          </button>
          {/* The full PivotTable Options dialog when the page provides one (2026-10-18); the small
              Totals-only panel below is the fallback. */}
          <button
            type="button"
            onClick={() => (onOpenPivotTableOptions ? (onClose(), onOpenPivotTableOptions()) : setPanel("pivotTableOptions"))}
            className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10"
          >
            PivotTable Options…
          </button>
          <button
            type="button"
            onClick={() => { onClose(); onHideFieldList(); }}
            className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10"
          >
            Hide Field List
          </button>
        </div>
      ) : panel === "sort" ? (
        <div className="flex flex-col gap-1">
          <button
            type="button"
            onClick={() => setPanel("root")}
            className="mb-1 inline-flex items-center gap-1 px-2 text-xs font-medium text-muted hover:text-ink"
          >
            <ChevronLeft className="h-3.5 w-3.5" /> Back
          </button>
          <button type="button" onClick={() => { setSort("Ascending"); onClose(); }} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
            Sort Smallest to Largest
          </button>
          <button type="button" onClick={() => { setSort("Descending"); onClose(); }} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
            Sort Largest to Smallest
          </button>
          <button type="button" onClick={() => setPanel("sortEditor")} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
            More Sort Options…
          </button>
        </div>
      ) : panel === "sortEditor" ? (
        // Excel opens the per-field "Sort (<field>)" dialog here too (2026-10-14, against a screenshot)
        // — for the cell's innermost Rows field (else its innermost Columns field).
        (() => {
          const realRows = request.rows.filter((f) => f !== MEASURES_PLACEHOLDER);
          const realColumns = request.columns.filter((f) => f !== MEASURES_PLACEHOLDER);
          const axis = realRows.length > 0 ? "Rows" : "Columns";
          const sortField = axis === "Rows" ? realRows[realRows.length - 1] : realColumns[realColumns.length - 1];
          if (!sortField) return <p className="p-2 text-sm text-muted">Add a Row or Column field to sort.</p>;
          const level = (axis === "Rows" ? request.rows : request.columns).indexOf(sortField);
          // The right-clicked cell's column is Excel's "selected column" for a Rows field (a total
          // column isn't one of the columns an item can be sorted by, so it offers none).
          const selectedItem = axis === "Rows" && !column.isTotal ? { labels: column.labels } : null;
          return <FieldSortDialog field={sortField} axis={axis} level={level} request={request} onChange={onChange} onBack={() => setPanel("sort")} onDone={onClose} selectedItem={selectedItem} />;
        })()
      ) : panel === "numberFormat" ? (
        <NumberFormatEditor format={format} onUpdateFormat={updateFormat} onBack={() => setPanel("root")} onDone={onClose} />
      ) : panel === "settings" ? (
        <ValueFieldSettingsPanel v={valueField} onUpdate={updateValue} onBack={() => setPanel("root")} onDone={onClose} rows={request.rows} columns={request.columns} />
      ) : panel === "pivotTableOptions" ? (
        <div className="space-y-3">
          <button
            type="button"
            onClick={() => setPanel("root")}
            className="mb-1 inline-flex items-center gap-1 px-2 text-xs font-medium text-muted hover:text-ink"
          >
            <ChevronLeft className="h-3.5 w-3.5" /> Back
          </button>
          {/* Same request.grandTotals/subtotals dropdowns as the Design ribbon's own Grand
              Totals/Subtotals buttons (PivotToolbar.jsx) — a second entry point onto the same live
              request, not a separate settings store, matching how every other control in this menu
              already works. Each a real Excel-shaped dropdown (4-way / 3-way), not a plain on/off
              checkbox (caught 2026-10-06, "عايز كل حاجة زي ال excel"). */}
          <div>
            <span className="mb-1 block text-xs font-medium text-muted">Grand Totals</span>
            <select
              className="w-full rounded-xl border border-slate-200 bg-white/90 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent"
              value={request.grandTotals}
              onChange={(e) => onChange({ ...request, grandTotals: e.target.value })}
            >
              <option value="Off">Off</option>
              <option value="RowsAndColumns">On for Rows and Columns</option>
              <option value="RowsOnly">On for Rows Only</option>
              <option value="ColumnsOnly">On for Columns Only</option>
            </select>
          </div>
          <div>
            <span className="mb-1 block text-xs font-medium text-muted">Subtotals</span>
            <select
              className="w-full rounded-xl border border-slate-200 bg-white/90 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent disabled:cursor-not-allowed disabled:opacity-40"
              value={request.subtotals}
              // Same Σ Values-isn't-a-real-field exclusion as PivotToolbar.jsx's own Subtotals
              // dropdown (caught 2026-10-06, full codebase review). A drilldown hierarchy has
              // subtotals of its own levels (Drilldown v2), so it enables this too.
              disabled={request.rows.filter((f) => f !== MEASURES_PLACEHOLDER).length < 2 && !request.rowsHierarchyDrilldown}
              onChange={(e) => onChange(applyReportSubtotals(request, e.target.value))}
            >
              <option value="Off">Do Not Show Subtotals</option>
              <option value="Bottom">Show all Subtotals at Bottom of Group</option>
              <option value="Top">Show all Subtotals at Top of Group</option>
            </select>
          </div>
        </div>
      ) : panel === "conditionalFormat" ? (
        <ConditionalFormatEditor
          format={valueField.conditionalFormat}
          onUpdate={(patch) => updateValue({ conditionalFormat: { ...(valueField.conditionalFormat ?? { type: "None" }), ...patch } })}
          onBack={() => setPanel("root")}
        />
      ) : (
        <div className="flex flex-col gap-1">
          <button
            type="button"
            onClick={() => setPanel("root")}
            className="mb-1 inline-flex items-center gap-1 px-2 text-xs font-medium text-muted hover:text-ink"
          >
            <ChevronLeft className="h-3.5 w-3.5" /> Back
          </button>
          {SHOW_VALUES_AS_OPTIONS.map((opt) => (
            <button
              key={opt.value}
              type="button"
              onClick={() => {
                // Both patches must land in ONE updateValue call — two separate onChange calls
                // would each build off the same stale `request` closure, so the second would
                // silently discard the first (caught during implementation, not live).
                const valuePatch = { showValuesAs: opt.value };
                // Modes that need a Base field get Excel's own defaults — the first Rows field and,
                // where there's a Base item, "(previous)" — instead of being sent without one (the
                // backend rejects that). Value Field Settings changes them.
                if (["PercentOf", "DifferenceFrom", "PercentDifferenceFrom", "PercentOfParentTotal"].includes(opt.value)) {
                  valuePatch.baseField = request.rows.find((f) => f !== MEASURES_PLACEHOLDER) ?? null;
                  valuePatch.baseItem = opt.value === "PercentOfParentTotal" ? "" : "(previous)";
                }
                // Running Total In / Rank: along the innermost Rows field, restarting for each outer
                // item (2026-10-20); with one Rows field that's the whole table, as before.
                if (["RunningTotal", "PercentRunningTotal", "Rank", "RankAscending"].includes(opt.value)) {
                  valuePatch.baseField = [...request.rows].reverse().find((f) => f !== MEASURES_PLACEHOLDER) ?? null;
                  valuePatch.baseItem = "";
                }
                const nextFormat = formatForShowValuesAs(format, valueField.showValuesAs, opt.value);
                if (nextFormat) valuePatch.format = nextFormat;
                const values = request.values.map((v, i) => (i === measureIndex ? { ...v, ...valuePatch } : v));
                // Grand Totals stay as they are (2026-10-20): the backend computes hidden ones for
                // % of Grand Total / Index, like Excel.
                onChange({ ...request, values });
                onClose();
              }}
              className={`rounded-lg px-2 py-1.5 text-left ${valueField.showValuesAs === opt.value ? "bg-gold/10 font-semibold" : "hover:bg-gold/10"}`}
            >
              {opt.label}
            </button>
          ))}
        </div>
      )}
    </ContextMenu>
  );
}

export default function PivotGrid({
  result,
  rows = [],
  columns = [],
  rowFieldLabels = [],
  columnFieldLabels = [],
  valueFields = [],
  dimensions = [],
  layout = "Tabular",
  repeatItemLabels = false,
  // Real Excel's per-field "Field Settings…" > Custom Name / Layout & Print (PLAN 6 Phase 2) — see
  // PivotBuilder.jsx's own fieldLayouts state WHY comment for why this is a sibling prop, not part
  // of `request`.
  fieldLayouts = [],
  // Real Excel's per-member OLAP drilldown (PLAN 8, 2026-10-13, EXPERIMENTAL) — { topLevelField,
  // expandedMemberUniqueNames } or null/undefined when not active.
  rowsHierarchyDrilldown = null,
  onToggleDrilldownMember = null,
  // (uniqueName, row) => boolean — whether that member is expanded on that row (per outer item for a
  // combined drilldown, Drilldown v3). Optional: without it, the drilldown's plain list is used.
  isDrilldownMemberExpanded = null,
  // The report Subtotals setting as applied to a drilldown hierarchy (see buildDisplayRows'
  // drilldownSubtotals). Passed explicitly rather than read off `request`, since a read-only
  // snapshot preview has no live request.
  drilldownSubtotals = "Top",
  // Columns drilldown (C4, 2026-10-14): its slot among `columns` (-1 = none) — header rows come from
  // buildColumnHeaderRows (parents spanning their children). The +/- callbacks get the member, the
  // header cell's first column (for its outer items) and the member's ancestors; a read-only
  // preview passes none, so no toggles are drawn.
  columnsDrillSlot = -1,
  isColumnDrilldownMemberExpanded = null,
  onToggleColumnDrilldownMember = null,
  request = null,
  onChange = null,
  onRefresh = null,
  onHideFieldList = null,
  // false when this grid is already nested inside another scrollable container (e.g. Modal.jsx's
  // own max-h-[90vh] overflow-auto card, used for the saved-snapshot interactive preview) — without
  // this, the grid's own 80vh scroll box nested inside the Modal's 90vh one produced two visible
  // scrollbars stacked on top of each other for the same content (caught live, 2026-10-07, "ليه في
  // 2 scroll bar؟"). The standalone PivotBuilder.jsx page (not inside any other scroll container)
  // still needs its OWN bounded box — see this div's own WHY comment below — so this defaults to
  // true there and is only turned off by the Modal-embedded caller.
  boundedHeight = true,
  // Excel's PivotTable Options > Layout & Format (2026-10-18) — empty/error cell text, merged labels,
  // report filter area layout. See utils/pivotOptions.js.
  pivotOptions = DEFAULT_PIVOT_OPTIONS,
  // Opens the PivotTable Options dialog (the value-cell menu's "PivotTable Options…"); null = none.
  onOpenPivotTableOptions = null,
  // The Builder's workspace layout (2026-10-20): the grid box flexes into its parent's leftover
  // height/width instead of hugging the table. The parent must be a flex column.
  fill = false,
  // Table zoom factor (1 = 100%), from the status bar.
  zoom = 1,
  // Called with the selected range's numbers ({ count, numericCount, sum, average }) or null when
  // nothing is selected — the status bar's Excel-style Sum/Average/Count.
  onSelectionStatsChange = null,
  // Number formats set on individual cells ({ [cellFormatKey]: format }, utils/cellFormats.js) — they
  // win over the value field's own format. onCellFormatsChange(next) receives Home › Number's edits.
  cellFormats = null,
  onCellFormatsChange = null,
  // { current } the grid fills with { applyNumberFormat(fn) } — how the ribbon formats the selection.
  commandRef = null,
  // A query is running: with no result yet the placeholder says so, instead of "choose fields"
  // (2026-10-20, found by the feature test — it showed while fields were already placed).
  loading = false,
}) {
  const [collapsed, setCollapsed] = useState(() => new Set());
  // The right-clicked value cell's context menu ({x, y, row, column, valueField, measureIndex}) and
  // the Show Details modal's own state (fetched separately, since the menu closes immediately on
  // click but the modal needs to stay open through the fetch and after) — see ValueCellContextMenu
  // above and the contextmenu handler / modal render below.
  const [cellMenu, setCellMenu] = useState(null);
  // The row-label right-click menu: { x, y, level, label, path, field, fieldLabel, collapseKey }.
  const [labelMenu, setLabelMenu] = useState(null);
  const [drillthrough, setDrillthrough] = useState(null);
  // A short message shown above the grid (e.g. double-clicking a cell with Show Details turned off),
  // cleared after a few seconds.
  const [notice, setNotice] = useState(null);
  useEffect(() => {
    if (!notice) return undefined;
    const timer = setTimeout(() => setNotice(null), 4000);
    return () => clearTimeout(timer);
  }, [notice]);
  // Keyboard navigation (Excel's own arrow-key cell selection, Enter to move down, Escape to
  // clear) — {row, col} are indices into `cells`/`rowHeaders`/`columnHeaders`, not screen position,
  // so the selection survives sorting/filtering as long as the same indices still exist. Only the
  // VALUE cells are navigable (not headers) — this app's grid isn't an editable spreadsheet, so the
  // one thing worth replicating is "move the active cell around with arrow keys to read values
  // without reaching for the mouse," not full worksheet-wide navigation.
  const [selectedCell, setSelectedCell] = useState(null);
  // The other corner of a selected RANGE (2026-10-20, Excel's Shift+click / Shift+arrow) — the range
  // runs from here to selectedCell; null/equal = just the one cell. Ctrl+C copies it (copyRange).
  const [anchorCell, setAnchorCell] = useState(null);
  // A fresh `result` (re-sort, re-filter, a field added/removed) can reshuffle which row/column
  // lands at a given index — `selectedCell` is plain {row, col} ARRAY INDICES with no identity
  // check against what's actually there now, so the ring highlight used to silently keep "selecting"
  // whatever row/column now happens to sit at the SAME indices in the new, unrelated result (caught
  // 2026-10-06, full codebase review). Cleared on every new result instead of carried forward.
  useEffect(() => {
    setSelectedCell(null);
    setAnchorCell(null);
  }, [result]);
  // Selects a cell; with Shift held the anchor stays put and the range grows, like Excel.
  const selectCell = (cell, extend = false) => {
    setSelectedCell(cell);
    if (!extend || !anchorCell) setAnchorCell(cell);
  };
  const gridRef = useRef(null);
  // True while the mouse button is held after pressing on a value cell (drag-to-select).
  const draggingRef = useRef(false);
  useEffect(() => {
    const stop = () => (draggingRef.current = false);
    document.addEventListener("mouseup", stop);
    return () => document.removeEventListener("mouseup", stop);
  }, []);
  // Same shape FieldPicker.jsx builds for FilterRow's own "choose a field"/"by measure" selects —
  // see FilterHeaderRow above, which reuses FilterRow unchanged for the Report Filter area's own
  // dropdown.
  const dimensionOptions = dimensions.map((d) => ({ value: d.field, label: d.displayName }));
  const selectedValueOptions = valueFields.filter((v) => v.field).map((v) => ({ value: v.field, label: v.field }));

  // Where (if either) Σ Values sits within rows/columns — rows/columns already contain it as a
  // plain element wherever it was placed (see utils/measuresPlacement.js), and rowFieldLabels/
  // columnFieldLabels already carry "Σ Values" at that same position too (PivotBuilder.jsx's own
  // fieldLabels helper maps the sentinel to that label) — no separate "append Σ Values" step
  // needed here anymore, unlike before Σ Values became positionable.
  const rowsMeasuresIndex = measuresIndexOf(rows);
  const columnsMeasuresIndex = measuresIndexOf(columns);
  const measuresInRows = rowsMeasuresIndex >= 0;
  const captionToIndex = useMemo(() => {
    if (!result || !valueFields?.length) return new Map();
    const headersOnMeasuresAxis = measuresInRows ? result.rowHeaders : result.columnHeaders;
    const measuresLevelIndex = measuresInRows ? rowsMeasuresIndex : columnsMeasuresIndex;
    return buildCaptionToIndex(headersOnMeasuresAxis, measuresLevelIndex, valueFields.length);
  }, [result, valueFields, measuresInRows, rowsMeasuresIndex, columnsMeasuresIndex]);
  const conditionalRanges = useMemo(() => {
    if (!result) return [];
    return measuresInRows
      ? result.rowHeaders.map((_, r) => detailRangeForRow(result, r))
      : result.columnHeaders.map((_, c) => detailRangeForColumn(result, c));
  }, [result, measuresInRows]);

  // Excel's "Show Details" — the clicked cell's row/column uniqueNames already include an entry at
  // the Σ Values level (a non-member placeholder caption, not a real dimension member — see
  // CubeQueryService.BuildHeaders), which must be stripped before sending, same as every other
  // place in this file that reads a header's REAL fields off rows/columns (e.g. captionToIndex
  // above already assumes the sentinel's own level isn't a real dimension level to group by).
  const runDrillthrough = async ({ row, column, valueField }) => {
    const rowUniqueNames = (row.uniqueNames ?? []).filter((_, i) => i !== rowsMeasuresIndex);
    const columnUniqueNames = (column.uniqueNames ?? []).filter((_, i) => i !== columnsMeasuresIndex);
    const label = valueField.customLabel || valueField.field;
    setDrillthrough({ loading: true, error: null, data: null, label });
    try {
      const data = await drillthroughPivot({
        rowUniqueNames,
        columnUniqueNames,
        measureField: valueField.field,
        // Only filters that narrow something — a field left at (All) is no filter (and the backend
        // rejects an empty member list), the same rule the pivot query uses.
        filters: request.filters.filter((f) => f.field && isFilterComplete(f)),
      });
      setDrillthrough({ loading: false, error: null, data, label });
    } catch (err) {
      setDrillthrough({ loading: false, error: err.message, data: null, label });
    }
  };

  // Show Details → Excel (2026-10-20). ExcelJS is loaded only when used (see PivotBuilder's export).
  const [drillthroughExporting, setDrillthroughExporting] = useState(false);
  const exportDrillthrough = async () => {
    if (!drillthrough?.data) return;
    setDrillthroughExporting(true);
    try {
      const { exportDetailsToExcel } = await import("../../utils/exportExcel");
      const safeLabel = drillthrough.label.replace(/[\\/:*?"<>|[\]]/g, " ").trim() || "Details";
      await exportDetailsToExcel({ columns: drillthrough.data.columns, rows: drillthrough.data.rows, filename: `Details - ${safeLabel}`, sheetName: "Details" });
    } catch (err) {
      setDrillthrough((current) => ({ ...current, exportError: err.message }));
    } finally {
      setDrillthroughExporting(false);
    }
  };

  // Excel never leaves the worksheet blank before the first field is placed — it shows this exact
  // placeholder inside the PivotTable's own bounds ("To build a report, choose fields from the
  // PivotTable Field List"). Matching that instead of rendering nothing (flagged live, 2026-10-04,
  // against a side-by-side screenshot of a brand-new real Excel PivotTable).
  // Slicers, timelines and the Report Filter area — drawn at the SAME place whether or not there's a
  // result (see both returns below), so React keeps them mounted while the pivot re-queries: they
  // used to vanish during every refresh and come back reset (a timeline switched to Years jumped
  // back to Months and re-loaded its items — 2026-10-21, found by the feature test). Excel keeps
  // them on screen while it refreshes.
  const filterWidgets = (
    <>
        {/* Excel's own Slicer control (Insert Slicer, PivotToolbar.jsx) — a separate button-grid
            widget bound to a field, living OUTSIDE the Report Filter area entirely even though it
            reads/writes the SAME request.filters entry (Members mode) a Report Filter row would.
            Rendered here instead (not as a Report Filter row below) for any field in
            request.slicerFields, which is also why that field is excluded from the Report Filter
            rows just below — Excel never shows both a Slicer AND a Report Filter row for the same
            field at once. */}
        {request && onChange && request.slicerFields.length > 0 && (
          <div className="mb-2 flex flex-wrap gap-2">
            {request.slicerFields.map((field) => {
              const filter = request.filters.find((f) => f.field === field) ?? null;
              return (
                <Slicer
                  key={field}
                  field={field}
                  displayName={dimensionOptions.find((d) => d.value === field)?.label ?? field}
                  filter={filter}
                  onChange={(nextFilter) => {
                    const next = request.filters.some((f) => f.field === field)
                      ? request.filters.map((f) => (f.field === field ? nextFilter : f))
                      : [...request.filters, nextFilter];
                    onChange({ ...request, filters: next });
                  }}
                  onRemove={() =>
                    onChange({
                      ...request,
                      slicerFields: request.slicerFields.filter((f) => f !== field),
                      filters: request.filters.filter((f) => f.field !== field),
                    })
                  }
                />
              );
            })}
          </div>
        )}
        {/* Excel's own Timeline control (Insert Timeline, PivotToolbar.jsx) — same relationship to
            request.filters as the Slicer block above, see Timeline.jsx's own WHY comment. */}
        {request && onChange && request.timelineFields.length > 0 && (
          <div className="mb-2 flex flex-wrap gap-2">
            {request.timelineFields.map((field) => {
              const filter = request.filters.find((f) => f.field === field) ?? null;
              return (
                <Timeline
                  key={field}
                  field={field}
                  displayName={dimensionOptions.find((d) => d.value === field)?.label ?? field}
                  filter={filter}
                  onChange={(nextFilter) => {
                    const next = request.filters.some((f) => f.field === field)
                      ? request.filters.map((f) => (f.field === field ? nextFilter : f))
                      : [...request.filters, nextFilter];
                    onChange({ ...request, filters: next });
                  }}
                  onRemove={() =>
                    onChange({
                      ...request,
                      timelineFields: request.timelineFields.filter((f) => f !== field),
                      filters: request.filters.filter((f) => f.field !== field),
                    })
                  }
                />
              );
            })}
          </div>
        )}
        {/* Excel's own Report Filter area sits ABOVE the PivotTable itself, not just in a side
            panel — see FilterHeaderRow's own WHY comment. Reads/writes the LIVE request.filters
            (not lastQuery), same as every other header-level Sort/Filter control in this file —
            still needs Run query to take effect, not a special case. A field already shown as a
            Slicer or Timeline widget above is excluded here — Excel never shows the same field as
            both a dedicated widget AND a Report Filter row at once. */}
        {/* Laid out by PivotTable Options > Layout & Format: Down, Then Over / Over, Then Down with N
            fields per column/row (filterAreaSlot — the same placement as the snapshot and export). */}
        {request && onChange && (() => {
          const shown = request.filters
            .map((filter, index) => ({ filter, index }))
            // Report Filter fields only — a Rows/Columns field's own filter is its header's funnel (2026-10-20).
            .filter(({ filter }) => !request.slicerFields.includes(filter.field) && !request.timelineFields.includes(filter.field) && !request.rows.includes(filter.field) && !request.columns.includes(filter.field));
          if (shown.length === 0) return null;
          return (
            <div className="mb-2 grid w-fit max-w-full gap-x-3 gap-y-1">
              {shown.map(({ filter, index }, position) => {
                const slot = filterAreaSlot(position, shown.length, pivotOptions);
                return (
                  <FilterHeaderRow
                    key={index}
                    filter={filter}
                    index={index}
                    request={request}
                    onChange={onChange}
                    dimensionOptions={dimensionOptions}
                    selectedValueOptions={selectedValueOptions}
                    style={{ gridColumn: slot.column + 1, gridRow: slot.row + 1 }}
                  />
                );
              })}
            </div>
          );
        })()}
    </>
  );

  if (!result) {
    return (
      <>
      {onSelectionStatsChange && <SelectionStatsReporter stats={null} onChange={onSelectionStatsChange} />}
      {/* Same slots as the main return: (command binder), filter widgets — see filterWidgets. */}
      {null}
      {filterWidgets}
      <div
        className={`flex min-h-[22rem] items-center justify-center rounded-lg border border-slate-200 bg-white ${fill ? "flex-1" : ""}`}
      >
        {/* Excel's own placeholder for an empty PivotTable, as a card on the empty sheet. */}
        <div className="max-w-sm rounded-xl border border-slate-200 bg-white px-8 py-7 text-center shadow-sm">
          <span className="mx-auto mb-3 flex h-11 w-11 items-center justify-center rounded-full bg-gold-soft">
            {loading ? <Loader2 className="h-5 w-5 animate-spin text-gold-strong" /> : <Table2 className="h-5 w-5 text-gold-strong" />}
          </span>
          {loading ? (
            <>
              <p className="text-sm font-semibold text-ink">Updating the PivotTable…</p>
              <p className="mt-1.5 text-xs leading-relaxed text-muted">Getting the numbers from the data server.</p>
            </>
          ) : (
            <>
              <p className="text-sm font-semibold text-ink">To build a report, choose fields from the PivotTable Fields list</p>
              <p className="mt-1.5 text-xs leading-relaxed text-muted">Tick a field, or drag it into Filters, Columns, Rows or Values.</p>
            </>
          )}
        </div>
      </div>
      </>
    );
  }

  const { rowHeaders, columnHeaders, cells } = result;
  // MAX across every row, not just row 0 — real Excel's per-member OLAP drilldown (PLAN 8) produces
  // genuinely ragged rows (a still-collapsed member has fewer labels than an expanded descendant),
  // so row 0 alone is no longer a reliable stand-in for "how many levels does this table have".
  // Harmless no-op for every non-drilldown pivot, where every row already shares the same depth.
  const rowDepth = rowHeaders.reduce((max, h) => Math.max(max, h.labels.length), 0);
  // One entry per row level, resolved once rather than re-looked-up per row/group — mirrors the
  // backend's HtmlSnapshotRenderer exactly (same field, same fallback-to-report-default order), so
  // the live grid and a saved snapshot never disagree on which fields share a column.
  const rowFieldLayouts = Array.from({ length: rowDepth }, (_, level) => fieldLayouts.find((f) => f.field === rows[level]) ?? null);
  const compactFlags = rowFieldLayouts.map((fl) => fl?.compact ?? layout === "Compact");
  // A drilldown hierarchy is always rendered as ONE shared indented column regardless of
  // reportLayout — matching real Excel's own natural-hierarchy rendering, and the only shape that
  // can represent ragged depth without a per-level field to anchor separate columns to (there is no
  // `rows[level]` field backing any level here — Rows is empty in this mode, see
  // PivotBuilder.jsx's own rowsHierarchyDrilldown WHY comment).
  //
  // Drilldown v3: a drilldown COMBINED with other Rows fields (its slot among `rows`) is laid out like
  // any Rows field instead — one label per field, the slot's own labels indented by member depth
  // (see buildDisplayRows' drillSlot). Only a drilldown ALONE keeps the single ragged column.
  const drillSlot = rowsHierarchyDrilldown ? rows.indexOf(rowsHierarchyDrilldown.topLevelField) : -1;
  const drilldownAlone = Boolean(rowsHierarchyDrilldown) && rows.filter((f) => f !== MEASURES_PLACEHOLDER && f !== rowsHierarchyDrilldown.topLevelField).length === 0;
  const combinedDrillSlot = rowsHierarchyDrilldown && !drilldownAlone ? drillSlot : -1;
  const rowHeaderGroups = drilldownAlone ? [{ start: 0, end: Math.max(0, rowDepth - 1) }] : buildRowHeaderGroups(rowDepth, compactFlags);
  const rowHeaderColumnCount = rowHeaderGroups.length;
  // Backend floor of 1 (not 0) whenever there's at least one row to label: with no Value fields at
  // all (Excel's own "removed the last Value field" state), Columns can legitimately come back with
  // ZERO headers (a literal empty MDX set — see MdxPivotQueryBuilder.BuildAxisFinal's own WHY
  // comment) — but the level-0 header row below is also where the ROW field's own title cell
  // (FieldHeaderMenu "Business Unit.BU Short Name" etc.) renders, via the `Array.from({length:
  // columnDepth})` loop a few lines down. columnDepth === 0 would skip that loop entirely and drop
  // the row headers' own title row along with the (correctly absent) column ones — caught live,
  // 2026-10-05, the first time this zero-Values case became reachable at all.
  // Columns drilldown (C2): a drilldown slot adds one header row per member depth — the shared
  // buildColumnHeaderRows (mirrored by the backend's ColumnHeaderLayout) decides rows and spans.
  const allColumnHeaderRows = buildColumnHeaderRows(columnHeaders, columnsDrillSlot, columnsMeasuresIndex);
  // One value field with Columns fields: Excel shows no row repeating the value's name under every
  // column — the name sits once in the top-left corner (2026-10-21). Only when Σ Values is the LAST
  // column level (its usual place, and always so with one value — PivotBuilder keeps it there), so
  // the Grand Total caption on the first row is never the one dropped.
  const labelLevels = columnHeaders[0]?.labels.length ?? 0;
  const hideSingleValueRow = valueFields.length === 1 && columnsDrillSlot < 0 && labelLevels > 1 && columnsMeasuresIndex === labelLevels - 1;
  const columnHeaderRows = hideSingleValueRow ? allColumnHeaderRows.filter((_, r) => r !== columnsMeasuresIndex) : allColumnHeaderRows;
  const singleValueCaption = hideSingleValueRow ? (columnHeaders.find((c) => !c.isTotal)?.labels[columnsMeasuresIndex] ?? valueFields[0]?.field ?? "") : null;
  const columnDepth = columnHeaders.length > 0 ? columnHeaderRows.length : rowHeaders.length > 0 ? 1 : 0;

  const toggleGroup = (key) => {
    setCollapsed((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  // The rows real Excel actually draws (item rows for compact-form levels, "ER Total"/"Grand Total"
  // captions, Top subtotals absorbed or moved) — see buildDisplayRows' own WHY comment. Shared with
  // exportExcel.js and mirrored by the backend's HtmlSnapshotRenderer, so the live grid, a saved
  // snapshot and an export all draw the same table. Collapsing an outermost item (Excel's +/-)
  // then replaces that item's rows with one row carrying its subtotal.
  const allDisplayRows = buildDisplayRows(rowHeaders, {
    rowHeaderGroups,
    rowFieldLayouts,
    repeatItemLabels,
    measuresLevel: rowsMeasuresIndex,
    drilldown: drilldownAlone,
    // Excel's Outline AND Compact layouts are both "outline form" (outer items on their own row);
    // only Tabular puts an item on its first child's row.
    outline: layout !== "Tabular",
    drilldownSubtotals,
    drillSlot: combinedDrillSlot,
  });
  // A drilldown hierarchy has its OWN per-member +/- (onToggleDrilldownMember) — the outer-item
  // collapse would be a second, conflicting toggle on the same rows.
  const collapsible = rowsHierarchyDrilldown ? new Set() : collapsibleKeys(allDisplayRows);
  const displayRows = applyCollapse(allDisplayRows, new Set([...collapsed].filter((k) => collapsible.has(k))), rowHeaderGroups);
  // Excel's "Expand/Collapse Entire Field" (2026-10-20): every item of one Rows level at once.
  const levelPrefix = (level) => `${level}\u0002`;
  const fieldCollapse = {
    canCollapse: (level) => [...collapsible].some((k) => k.startsWith(levelPrefix(level))),
    isCollapsed: (level) => [...collapsed].some((k) => k.startsWith(levelPrefix(level))),
    collapse: (level) => setCollapsed((prev) => new Set([...prev, ...[...collapsible].filter((k) => k.startsWith(levelPrefix(level)))])),
    expand: (level) => setCollapsed((prev) => new Set([...prev].filter((k) => !k.startsWith(levelPrefix(level))))),
  };

  // Excel's Move "X" Up/Down/To Beginning/To End, and dragging an item onto another (2026-10-20):
  // the item moves among its SIBLINGS (same outer items), and the level gets a manual order
  // (LevelSort.CustomOrder) — the level's labels in display order with the siblings rearranged.
  const moveItem = (level, path, where) => {
    if (!request || !onChange) return;
    const label = path[level];
    const parent = path.slice(0, level);
    const unique = (labels) => [...new Set(labels)];
    const atLevel = allDisplayRows.filter((d) => d.kind !== "grand" && d.path.length > level);
    const siblings = unique(atLevel.filter((d) => parent.every((p, k) => d.path[k] === p)).map((d) => d.path[level]));
    const from = siblings.indexOf(label);
    if (from < 0) return;
    const reordered = siblings.filter((s) => s !== label);
    const to =
      where === "up" ? Math.max(0, from - 1)
      : where === "down" ? Math.min(reordered.length, from + 1)
      : where === "beginning" ? 0
      : where === "end" ? reordered.length
      : reordered.indexOf(where.before);
    if (to < 0) return;
    reordered.splice(to, 0, label);
    const siblingSet = new Set(siblings);
    let next = 0;
    const order = unique(atLevel.map((d) => d.path[level])).map((l) => (siblingSet.has(l) ? reordered[next++] : l));
    onChange({
      ...request,
      sort: request.sort?.axis === "Rows" ? null : request.sort,
      rowLevelSorts: [...request.rowLevelSorts.filter((s) => s.level !== level), { level, direction: "Ascending", byMeasureField: null, customOrder: order }],
    });
  };

  // Right-click and drag handlers for one row label (a real Rows field's item — not a total row,
  // Σ Values, or a drilldown member, which have their own controls).
  const labelHandlers = (level, displayRow) => {
    const field = rows[level];
    if (!request || !onChange || !field || field === MEASURES_PLACEHOLDER || drilldownAlone || level === combinedDrillSlot) return {};
    if (displayRow.kind === "subtotal" || displayRow.kind === "grand" || displayRow.path.length <= level) return {};
    const label = displayRow.path[level];
    const parentKey = JSON.stringify(displayRow.path.slice(0, level));
    // A group's own label, or a date/number bucket ("Q1 2025", "0-10"), isn't a cube member — Keep
    // Only/Hide would filter on a name the cube doesn't have (found in review 2026-10-20).
    const filterable =
      !(request.itemGroups ?? []).some((g) => g.field === field && g.name === label) &&
      !request.dateGroupings.some((g) => g.field === field) &&
      !request.numericGroupings.some((g) => g.field === field);
    return {
      onContextMenu: (e) => {
        e.preventDefault();
        setLabelMenu({
          x: e.clientX,
          y: e.clientY,
          level,
          label,
          filterable,
          path: displayRow.path,
          field,
          fieldLabel: rowFieldLayouts[level]?.customName || rowFieldLabels[level] || field,
          collapseKey: collapseKey(level, displayRow.path),
        });
      },
      draggable: true,
      onDragStart: (e) => {
        e.dataTransfer.setData(GRID_ITEM_DRAG_MIME, JSON.stringify({ level, parentKey, label }));
        e.dataTransfer.effectAllowed = "move";
      },
      onDragOver: (e) => {
        if (e.dataTransfer.types.includes(GRID_ITEM_DRAG_MIME)) e.preventDefault();
      },
      onDrop: (e) => {
        const raw = e.dataTransfer.getData(GRID_ITEM_DRAG_MIME);
        if (!raw) return;
        e.preventDefault();
        e.stopPropagation();
        const dragged = JSON.parse(raw);
        // Only within the same field and the same outer item, like Excel.
        if (dragged.level !== level || dragged.parentKey !== parentKey || dragged.label === label) return;
        moveItem(level, [...displayRow.path.slice(0, level), dragged.label], { before: label });
      },
    };
  };
  // PivotTable Options > Display (2026-10-19): "Show expand/collapse buttons" hides every +/- (an
  // outer item's, a drilldown member's, a Columns drilldown header's); "Display field captions and
  // filter drop downs" leaves the field-name header cells empty — no caption, no arrow — like Excel.
  const showExpandButtons = pivotOptions.showExpandButtons !== false;
  const FieldCaption = pivotOptions.showFieldCaptions === false ? NoCaption : FieldHeaderMenu;
  // "Merge and center cells with labels": an item's label cell spans its rows (computeLabelSpans).
  const labelSpans = pivotOptions.mergeLabels ? computeLabelSpans(displayRows, rowHeaderGroups, { drillSlot: combinedDrillSlot, drilldown: drilldownAlone }) : null;
  // Value-carrying rows in display order — what arrow-key navigation steps through (an item row
  // with no subtotal has no value cells to select).
  const navigableRows = displayRows.filter((d) => d.sourceIndex !== null).map((d) => d.sourceIndex);

  // Grid lines on every header/body cell (border on all four sides, not just a per-row bottom
  // border) — the single biggest visual gap vs a real spreadsheet grid (flagged live, 2026-09-28).
  // Lightened from slate-300 to slate-200 2026-10-12 ("شكله قذر") — full-strength gridlines on
  // every single cell edge read as visually busy/heavy at this app's own font size and padding;
  // the thinner, lighter tone keeps the grid legible (still a real bordered cell, not borderless)
  // without it dominating the page the way a literal flat worksheet's gridlines do.
  const cellBorder = "border border-slate-200";
  // Header band (2026-10-20 UI refresh): a flat slate-100 band with slate-300 lines and ink text,
  // closed off by a 2px gold rule under the last header row (the thead's own class below) — the app's
  // surfaces went flat slate with gold as the one accent, so the header follows the same material.
  // It replaces the gold gradient band (2026-10-12), whose white/muted text kept losing contrast.
  // (A solid dark-ink band was tried and rejected 2026-10-12 as foreign to the app.)
  const headerCellBorder = "border border-slate-300";
  const headerBg = "bg-slate-100";

  // Excel's own worksheet row-number gutter (the gray "1, 2, 3…" column down the left edge of
  // every sheet, outside the PivotTable report itself) — requested live 2026-10-12 ("ليه مافيش ال
  // row number زي ال excel"). One number per literal rendered table row (header rows included, same
  // as a real worksheet counting from row 1 at the very top), not per pivot data row — a header row
  // is still a real worksheet row. headerRowCount below is how many header <tr>s render before the
  // body: the optional column-field-label row (only when columnFieldLabels is non-empty) plus one
  // row per column header level.
  const headerRowCount = (columnFieldLabels.length > 0 ? 1 : 0) + columnDepth;
  // Two variants: the gutter's header-row cells share the header band; its body cells keep the
  // light-gray worksheet-gutter look.
  const rowNumberHeaderClass = `${headerCellBorder} ${headerBg} px-2 py-2 text-center text-xs font-medium text-slate-500`;
  const rowNumberBodyClass = `${cellBorder} bg-slate-50 px-2 py-1.5 text-center text-xs font-medium text-slate-400`;

  // Excel's own arrow-key cell navigation — see selectedCell's own WHY comment above. Skips over
  // collapsed/hidden rows the same way the mouse already does (clicking a collapsed group's hidden
  // detail rows isn't possible either, since they're not rendered at all).
  const handleGridKeyDown = (e) => {
    // Keys typed in a popover/menu (portaled, but React still bubbles them here) or in any input
    // belong to that control — Ctrl+C in a filter search box must copy its text, not the cell range.
    const target = e.target;
    if (!e.currentTarget.contains(target) || target.closest?.("input, textarea, select, [contenteditable='true']")) return;
    // Ctrl+C / ⌘C copies the selected range, like Excel.
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "c" && selectedCell) {
      e.preventDefault();
      copyRange();
      return;
    }
    if (!["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight", "Enter", "Escape"].includes(e.key)) return;
    e.preventDefault();
    if (e.key === "Escape") {
      setSelectedCell(null);
      setAnchorCell(null);
      return;
    }
    if (!selectedCell) {
      if (navigableRows.length > 0 && columnHeaders.length > 0) selectCell({ row: navigableRows[0], col: 0 });
      return;
    }
    const { row, col } = selectedCell;
    const extend = e.shiftKey && e.key !== "Enter";
    if (e.key === "ArrowLeft") {
      selectCell({ row, col: Math.max(0, col - 1) }, extend);
    } else if (e.key === "ArrowRight") {
      selectCell({ row, col: Math.min(columnHeaders.length - 1, col + 1) }, extend);
    } else if (e.key === "ArrowUp" || e.key === "ArrowDown" || e.key === "Enter") {
      // Steps in DISPLAY order (a Top subtotal can be drawn after its item's rows in Tabular, so
      // source order isn't screen order).
      const position = navigableRows.indexOf(row);
      const next = navigableRows[position + (e.key === "ArrowUp" ? -1 : 1)];
      if (position >= 0 && next !== undefined) selectCell({ row: next, col }, extend);
    }
  };

  // The selected range in DISPLAY order: row positions in navigableRows and a column span.
  const rangeBounds = () => {
    if (!selectedCell) return null;
    const anchor = anchorCell ?? selectedCell;
    const a = navigableRows.indexOf(anchor.row);
    const b = navigableRows.indexOf(selectedCell.row);
    if (a < 0 || b < 0) return null;
    return { top: Math.min(a, b), bottom: Math.max(a, b), left: Math.min(anchor.col, selectedCell.col), right: Math.max(anchor.col, selectedCell.col) };
  };
  const bounds = rangeBounds();
  const isMultiCellRange = bounds && (bounds.top !== bounds.bottom || bounds.left !== bounds.right);
  const inRange = (rowIndex, colIndex) => {
    if (!isMultiCellRange) return false;
    const position = navigableRows.indexOf(rowIndex);
    return position >= bounds.top && position <= bounds.bottom && colIndex >= bounds.left && colIndex <= bounds.right;
  };
  // The selected range's Sum/Average/Count for the status bar (SelectionStatsReporter below).
  const selectedRows = bounds ? navigableRows.slice(bounds.top, bounds.bottom + 1) : [];
  const selectedValues = selectedRows.flatMap((r) => (cells[r] ?? []).slice(bounds.left, bounds.right + 1));
  const baseStats = bounds ? selectionStats(selectedValues) : null;
  // Plus what Home › Number needs (2026-10-20): the value fields the selection covers (a format
  // applies to a whole value field, as in an Excel PivotTable) and the most decimals its numbers
  // show in General (up to 2) — where Increase/Decrease Decimal starts from.
  const rangeStats = baseStats && {
    ...baseStats,
    measures: [
      ...new Set(
        selectedRows.flatMap((r) =>
          Array.from({ length: bounds.right - bounds.left + 1 }, (_, k) =>
            measuresInRows ? measureIndexOf(rowHeaders[r], rowsMeasuresIndex, captionToIndex) : measureIndexOf(columnHeaders[bounds.left + k], columnsMeasuresIndex, captionToIndex),
          ),
        ),
      ),
    ]
      .filter((m) => m !== null)
      .sort((a, b) => a - b),
    generalDecimals: Math.max(0, ...selectedValues.filter((v) => typeof v === "number" && Number.isFinite(v)).map((v) => (String(Math.round(v * 100) / 100).split(".")[1] ?? "").length)),
    // The format type all selected cells share (the Number Format dropdown shows it), or null.
    formatType: sharedFormatType(
      selectedRows.flatMap((r) =>
        Array.from({ length: bounds.right - bounds.left + 1 }, (_, k) => {
          const c = bounds.left + k;
          const m = measuresInRows ? measureIndexOf(rowHeaders[r], rowsMeasuresIndex, captionToIndex) : measureIndexOf(columnHeaders[c], columnsMeasuresIndex, captionToIndex);
          return cellFormats?.[cellFormatKey(rowHeaders[r], columnHeaders[c])] ?? formatForColumn(m, valueFields);
        }),
      ),
    ),
  };

  // A cell's format: its own (Home › Number on a selection) or its value field's.
  const formatOfCell = (rowIndex, colIndex, measureIndex) => cellFormats?.[cellFormatKey(rowHeaders[rowIndex], columnHeaders[colIndex])] ?? formatForColumn(measureIndex, valueFields);
  const measureOfCell = (rowIndex, colIndex) =>
    measuresInRows ? measureIndexOf(rowHeaders[rowIndex], rowsMeasuresIndex, captionToIndex) : measureIndexOf(columnHeaders[colIndex], columnsMeasuresIndex, captionToIndex);
  // The decimals a number shows in General (up to 2) — where Increase/Decrease Decimal starts.
  const generalDecimalsOf = (v) => (typeof v === "number" && Number.isFinite(v) ? (String(Math.round(v * 100) / 100).split(".")[1] ?? "").length : 2);
  const selectedCellList = bounds
    ? selectedRows.flatMap((r) => Array.from({ length: bounds.right - bounds.left + 1 }, (_, k) => ({ row: r, col: bounds.left + k })))
    : [];
  // Handed to the ribbon through commandRef (CommandBinder, below).
  const gridCommands = {
    // Home › Number: fn(format, shownDecimals) → next format, for every selected cell.
    applyNumberFormat: (fn) => {
      if (!onCellFormatsChange || selectedCellList.length === 0) return;
      const next = { ...(cellFormats ?? {}) };
      for (const { row, col } of selectedCellList) {
        next[cellFormatKey(rowHeaders[row], columnHeaders[col])] = fn(formatOfCell(row, col, measureOfCell(row, col)) ?? undefined, generalDecimalsOf(cells[row]?.[col]));
      }
      onCellFormatsChange(next);
    },
  };

  // Clicking a column header selects its whole column (every value row), like clicking a column in
  // Excel; Shift+click extends from the current anchor column (2026-10-20).
  const selectColumns = (first, last, extend) => {
    if (navigableRows.length === 0) return;
    const top = navigableRows[0];
    const bottom = navigableRows[navigableRows.length - 1];
    if (extend && anchorCell) {
      const target = last >= anchorCell.col ? last : first;
      setAnchorCell({ row: top, col: anchorCell.col });
      setSelectedCell({ row: bottom, col: target });
    } else {
      setAnchorCell({ row: top, col: first });
      setSelectedCell({ row: bottom, col: last });
    }
    gridRef.current?.focus();
  };
  const columnSelected = (first, last) => Boolean(bounds) && bounds.top === 0 && bounds.bottom === navigableRows.length - 1 && first >= bounds.left && last <= bounds.right;

  // Copies the range as tab-separated text (what Excel puts on the clipboard), each value in its
  // displayed format; one row of the range per line.
  const copyRange = async () => {
    if (!bounds) return;
    const lines = navigableRows.slice(bounds.top, bounds.bottom + 1).map((rowIndex) =>
      Array.from({ length: bounds.right - bounds.left + 1 }, (_, k) => {
        const colIndex = bounds.left + k;
        const measureIndex = measuresInRows
          ? measureIndexOf(rowHeaders[rowIndex], rowsMeasuresIndex, captionToIndex)
          : measureIndexOf(columnHeaders[colIndex], columnsMeasuresIndex, captionToIndex);
        return cellText(cells[rowIndex]?.[colIndex], formatOfCell(rowIndex, colIndex, measureIndex), pivotOptions);
      }).join("\t"),
    );
    try {
      await navigator.clipboard.writeText(lines.join("\n"));
      const count = lines.length * (bounds.right - bounds.left + 1);
      setNotice(`Copied ${count} cell${count === 1 ? "" : "s"}.`);
    } catch {
      setNotice("Couldn't copy — the browser blocked clipboard access.");
    }
  };

  return (
    <>
      {onSelectionStatsChange && <SelectionStatsReporter stats={rangeStats} onChange={onSelectionStatsChange} />}
      {commandRef && <CommandBinder target={commandRef} value={gridCommands} />}
      {filterWidgets}
      {notice && (
        <p role="status" className="mb-2 w-fit rounded-lg border border-amber-300 bg-amber-50 px-3 py-1.5 text-xs text-amber-800">
          {notice}
        </p>
      )}
      {/* ONE element now does both the "card" chrome AND the scrolling — previously these were two
          nested divs (a decorative outer `card` with no overflow set, wrapping this scrolling box),
          which is exactly why the table had no visible border-radius ("الجدول مالهوش border radius",
          2026-10-12): rounding the OUTER div does nothing unless something actually CLIPS the
          content to that rounded shape, and putting overflow-hidden on that outer div (tried
          earlier, 2026-10-06) broke the thead's sticky top-0 instead — ANY ancestor with overflow
          other than visible counts as a scroll container for position:sticky purposes, so a second,
          non-scrolling overflow box ABOVE the real scrolling one made the browser stick the header
          to the wrong (non-scrolling) box. Merging the two divs sidesteps the conflict entirely:
          overflow-auto right here already clips its own content to ITS OWN border-radius (true of
          any overflow value other than visible, not just hidden), and it's still the exact same,
          single element the thead's sticky top-0 was always measuring against — no second ancestor
          in the chain to confuse it. w-fit still hugs the table's real content width instead of
          stretching full-page (same reasoning as before this merge). */}
      <div
        ref={gridRef}
        // bg/border/rounded-lg: the app's flat panel treatment (index.css `.glass-panel`), applied
        // directly here instead of via the class so the SAME element also owns the overflow (see
        // this div's own WHY comment above). custom-scrollbar (index.css) swaps in the app's thin
        // scrollbar.
        // fill (the Builder's workspace, 2026-10-20): the box takes all the space its flex-column
        // parent leaves — a white "worksheet" the table sits in at the top-left, scrolling inside
        // itself — instead of hugging the table (w-fit) under an 80vh cap, which left a big empty
        // band beside/below a small pivot.
        // No backdrop-blur-xl here — this element is BOTH the
        // overflow-auto scroll container AND the ancestor of the <thead>'s position:sticky (a few
        // lines down), and backdrop-filter on an element that's also a scrolling ancestor of a
        // sticky descendant is a known Chromium compositing bug: the browser has to re-composite the
        // blurred backdrop together with the sticky repositioning on every scroll frame, and during
        // a fast scroll it visibly falls behind, letting scrolled-past rows render ON TOP of the
        // header for the whole gesture (confirmed live, 2026-10-12 — persisted even after fixing a
        // real but separate z-index bug AND adding will-change-transform, which is the normal first
        // fix for sticky-during-scroll glitches but doesn't address a backdrop-filter-specific one).
        // A custom-drawn vertical-only scrollbar thumb (starting below the header instead of at its
        // top) was also tried here 2026-10-12 to fix the native thumb's travel range including the
        // header's own height — reverted the same day: hiding just the native VERTICAL scrollbar
        // relies on the WebKit-only `::-webkit-scrollbar:vertical` selector, which this app's actual
        // runtime didn't honor, so the native one stayed visible right alongside the new custom thumb
        // ("بقا في اتنين scroll bars"). Real fix needs either a cross-browser-safe hiding technique
        // or splitting header/body into genuinely separate scroll regions with synced column widths —
        // both bigger, separately-scoped tasks. boundedHeight=false (Modal.jsx's own preview) has no
        // sticky header to protect (see the thead's own WHY comment below) and no independent
        // scrolling of its own either, so overflow-hidden there is just the plain corner-clipping.
        className={`rounded-lg border border-slate-200 bg-white shadow-sm custom-scrollbar outline-none ${
          fill ? "min-h-[16rem] w-full flex-1 overflow-auto xl:min-h-0" : `w-fit max-w-full ${boundedHeight ? "max-h-[80vh] overflow-auto" : "overflow-hidden"}`
        }`}
        tabIndex={0}
        onKeyDown={handleGridKeyDown}
      >
        <table
          // Excel's sheet font (Segoe UI on Windows) for the worksheet itself (2026-10-20), so the
          // numbers read like a spreadsheet; the app's chrome keeps the HMIS typeface.
          className="select-none border-collapse bg-white font-['Segoe_UI_Variable_Text','Segoe_UI',system-ui,sans-serif] text-sm"
          // View zoom (the status bar's zoom control, 2026-10-20) — CSS zoom re-lays the table out
          // at the new size, so the scroll box's scrollbars track it (a transform wouldn't).
          style={zoom !== 1 ? { zoom } : undefined}
          // PivotTable Options › Alt Text: the table's accessible name and description.
          aria-label={pivotOptions.altTextTitle || pivotOptions.name || undefined}
          aria-description={pivotOptions.altTextDescription || undefined}
          title={pivotOptions.altTextTitle || undefined}
        >
          {/* Sticky only in the standalone, boundedHeight case — it relies on gridRef's OWN
              dedicated overflow-auto box right above being the nearest scrolling ancestor (see its
              own WHY comment). Inside Modal.jsx's preview (boundedHeight=false), the scroll
              container is the Modal's shared outer box instead, with sibling content (the title,
              the Export button) ABOVE the table in that SAME scrolling box — a <thead> stuck via
              position:sticky there rendered with rows visibly bleeding/ghosting through it instead
              of a clean opaque header (caught live, 2026-10-07, "الشكل مش حلو ان ال rows بتبان من
              فوقيه"). Simplest reliable fix: don't stick it there at all — a small preview popup
              scrolling past its own header once is an acceptable tradeoff for not fighting a
              cross-browser sticky-table-header-in-a-shared-scroll-container quirk. */}
          {/* z-30, not z-10: a SELECTED value cell below uses z-20 (so its ring renders above
              neighboring cells' own borders in this border-collapse table) — with the header at only
              z-10, scrolling a selected cell up underneath the sticky header let it win the stacking
              order and render ON TOP of the header instead of being covered by it. The header must
              always out-rank anything that can appear in the body.
              will-change-transform: a SEPARATE bug from the z-index one above, caught right after it
              — confirmed live (2026-10-12) that a scrolled-past row shows through the header for the
              ENTIRE DURATION of an active scroll gesture (not just a brief one-frame flicker, and not
              tied to any selected cell), only during the scroll itself. This is a known Chromium
              compositing quirk: gridRef (this table's scrolling ancestor, a few lines up) has its own
              backdrop-blur-xl for the app's glass-panel look, and a `backdrop-filter` on an ancestor
              can make a descendant position:sticky element's repaint lag behind the actual scroll
              during fast scroll frames, since the browser has to re-composite the blurred backdrop
              and the sticky repositioning together. will-change: transform (Tailwind's
              `will-change-transform`) hints the browser to promote the header to its own compositing
              layer up front, instead of only reacting once scrolling starts — the standard fix for
              this exact class of "sticky element flickers/lags during scroll" bug. */}
          {/* [&>tr:last-child>th]: the 2px gold rule closing off the header band (see headerBg). */}
          <thead className={`[&>tr:last-child>th]:border-b-2 [&>tr:last-child>th]:border-b-gold ${boundedHeight || fill ? "sticky top-0 z-30 will-change-transform" : ""}`}>
            {columnFieldLabels.length > 0 && (
              <tr>
                <th className={rowNumberHeaderClass}>1</th>
                <th colSpan={rowHeaderColumnCount} className={`${headerCellBorder} ${headerBg} px-3 py-2 text-left`}>
                  {singleValueCaption && <span className="text-sm font-semibold text-ink">{singleValueCaption}</span>}
                </th>
                {/* Math.max(1, ...) — columnHeaders can legitimately be EMPTY (a Columns field
                    configured but the last Value field removed, see columnDepth's own WHY comment
                    above) while a Columns field is still configured (columnFieldLabels.length > 0
                    still true) — colSpan={0} isn't a meaningful table attribute and collapsed this
                    cell (with its own working FieldHeaderMenu dropdown) unpredictably instead of
                    just rendering at its practical minimum width (caught 2026-10-06, full codebase
                    review). */}
                <th colSpan={Math.max(1, columnHeaders.length)} className={`${headerCellBorder} border-b-2 border-b-slate-300 ${headerBg} px-2 py-2`}>
                  <FieldCaption
                    label={columnFieldLabels.join(" › ")}
                    // A Columns drilldown sorts from here (A→Z/Z→A or by value), whatever else is on
                    // Columns — the backend counts it as the axis's one sortable field. Sort only,
                    // same as the Rows drilldown's "Row Labels" header.
                    field={columnsDrillSlot >= 0 ? columns[columnsDrillSlot] : null}
                    // Otherwise Excel's "Column Labels" arrow: every real Columns field, picked from
                    // its "Select field:" box when there are several. Σ Values is left out — it used
                    // to count as a field here, so one field + Σ Values got no arrow at all (2026-10-21).
                    fieldChoices={
                      columnsDrillSlot >= 0
                        ? null
                        : columns
                            .map((f, l) => ({ field: f, level: l, label: columnFieldLabels[l] || f }))
                            .filter((c) => c.field && c.field !== MEASURES_PLACEHOLDER)
                    }
                    axis="Columns"
                    request={request}
                    onChange={onChange}
                    muted
                    sortOnly={columnsDrillSlot >= 0}
                  />
                </th>
              </tr>
            )}
            {Array.from({ length: columnDepth }).map((_, level) => (
              <tr key={level}>
                <th className={rowNumberHeaderClass}>{(columnFieldLabels.length > 0 ? 1 : 0) + level + 1}</th>
                {level === 0 && (
                  <>
                    {rowHeaderGroups.map((group) =>
                      // A drilldown hierarchy ALWAYS renders as the merged "Row Labels" caption
                      // below, even when it currently only has 1 level (nothing expanded yet) — a
                      // real field, degenerate OR not, is the ONLY thing FieldHeaderMenu's own
                      // `field` can meaningfully point at, and Rows is empty in drilldown mode (see
                      // rowsHierarchyDrilldown's own WHY comment), so this branch would otherwise
                      // render a blank header (caught live, 2026-10-13, "الheader...مش باين" — this
                      // was actually the SAME degenerate-depth-1 case as the color-contrast bug, not
                      // a second contrast issue).
                      group.start === group.end && !drilldownAlone ? (
                        <th key={group.start} rowSpan={columnDepth} className={`${headerCellBorder} border-b-2 border-b-gold ${headerBg} px-2 py-2`}>
                          <FieldCaption
                            label={rowFieldLayouts[group.start]?.customName || rowFieldLabels[group.start] || ""}
                            field={rows[group.start] ?? null}
                            axis="Rows"
                            level={group.start}
                            request={request}
                            onChange={onChange}
                            muted
                            // A combined drilldown slot's own column: sort only (Drilldown v3).
                            sortOnly={group.start === combinedDrillSlot}
                            fieldCollapse={showExpandButtons ? fieldCollapse : null}
                          />
                        </th>
                      ) : (
                        // A merged (compact-style) column spanning 2+ fields — not one-to-one with a
                        // single field, so this stays a plain (non-interactive) header, matching real
                        // Excel's own Compact form literal "Row Labels" caption (confirmed live
                        // 2026-10-05 against a real Excel screenshot) for BOTH the old all-fields-
                        // merged case and a new partial merge — there's no confirmed screenshot of
                        // what a partial merge's own header shows, so this reuses the one convention
                        // already verified for the full-merge case rather than guessing a new one.
                        <th key={group.start} rowSpan={columnDepth} className={`${headerCellBorder} border-b-2 border-b-gold ${headerBg} px-2 py-2`}>
                          {drilldownAlone ? (
                            // A drilldown hierarchy sorts from here (Drilldown v2) — sort only, see
                            // FieldHeaderMenu's sortOnly WHY comment.
                            <FieldCaption
                              label="Row Labels"
                              field={rowsHierarchyDrilldown.topLevelField ?? null}
                              axis="Rows"
                              request={request}
                              onChange={onChange}
                              muted
                              sortOnly
                            />
                          ) : (
                            // Excel's Compact "Row Labels" arrow: sort/filter any of the fields
                            // sharing this column, picked from its "Select field:" box (Σ Values
                            // and a drilldown slot have no menu of their own, so they're left out).
                            <FieldCaption
                              label="Row Labels"
                              axis="Rows"
                              request={request}
                              onChange={onChange}
                              muted
                              fieldChoices={Array.from({ length: group.end - group.start + 1 }, (_, i) => group.start + i)
                                .filter((l) => rows[l] && rows[l] !== MEASURES_PLACEHOLDER && l !== combinedDrillSlot)
                                .map((l) => ({ field: rows[l], level: l, label: rowFieldLayouts[l]?.customName || rowFieldLabels[l] || rows[l] }))}
                              fieldCollapse={showExpandButtons ? fieldCollapse : null}
                            />
                          )}
                        </th>
                      ),
                    )}
                  </>
                )}
                {/* A Grand Total column shows its caption once, in the top header row, like Excel; an
                    expanded Columns drilldown member spans its children (buildColumnHeaderRows). */}
                {(columnHeaderRows[level] ?? []).map((cell) => {
                  const column = columnHeaders[cell.sourceIndex];
                  const uniqueName = cell.memberUniqueName && cell.memberUniqueName.endsWith("]") ? cell.memberUniqueName : null;
                  const toggle = uniqueName && onToggleColumnDrilldownMember && showExpandButtons;
                  // A member's ancestors: those of the column it heads, down to the member itself.
                  const allAncestors = column?.drill?.ancestorUniqueNames ?? [];
                  const ownIndex = allAncestors.indexOf(uniqueName);
                  const ancestors = ownIndex >= 0 ? allAncestors.slice(0, ownIndex) : allAncestors;
                  const expanded = toggle ? (isColumnDrilldownMemberExpanded?.(uniqueName, column) ?? false) : false;
                  const lastCol = cell.sourceIndex + Math.max(cell.span, 1) - 1;
                  return (
                    <th
                      key={cell.sourceIndex}
                      colSpan={cell.span > 1 ? cell.span : undefined}
                      className={`${headerCellBorder} ${columnSelected(cell.sourceIndex, lastCol) ? "bg-gold-soft" : headerBg} cursor-pointer px-2 py-2 hover:bg-slate-200/70`}
                      title="Click to select this column — Shift+click to extend"
                      onClick={(e) => {
                        if (e.target.closest("button")) return; // the drilldown +/- toggle
                        selectColumns(cell.sourceIndex, lastCol, e.shiftKey);
                      }}
                    >
                      {toggle ? (
                        <span className={`flex items-center gap-1 ${cell.span > 1 && !pivotOptions.mergeLabels ? "justify-start" : "justify-center"}`}>
                          <button
                            type="button"
                            onClick={() => onToggleColumnDrilldownMember(uniqueName, column, ancestors)}
                            className="text-slate-500 hover:text-ink"
                            aria-label={expanded ? "Collapse" : "Expand"}
                          >
                            {expanded ? <ChevronDown className="h-3 w-3" /> : <ChevronRight className="h-3 w-3" />}
                          </button>
                          <HeaderCell isTotal={cell.isTotal} alignLeft={cell.span > 1 && !pivotOptions.mergeLabels}>{cell.text}</HeaderCell>
                        </span>
                      ) : (
                        <HeaderCell isTotal={cell.isTotal} alignLeft={cell.span > 1 && !pivotOptions.mergeLabels}>{cell.text}</HeaderCell>
                      )}
                    </th>
                  );
                })}
              </tr>
            ))}
          </thead>
          <tbody>
            {displayRows.map((displayRow, displayIndex) => {
              // The result row whose VALUES this display row shows — null for an item row with no
              // subtotal of its own (blank value cells, like Excel).
              const rowIndex = displayRow.sourceIndex;
              const row = rowIndex !== null ? rowHeaders[rowIndex] : null;
              // The member this row LABELS — differs from `row` only for a drilldown parent whose
              // values were blanked by its Subtotals setting; its +/- toggle still needs its identity.
              const labelRow = displayRow.labelSourceIndex !== undefined ? rowHeaders[displayRow.labelSourceIndex] : row;
              const isTotalRow = displayRow.kind === "subtotal" || displayRow.kind === "grand" || (displayRow.kind === "header" && row !== null);
              // Outer-item rows (compact form) and collapsed items get a bold label like Excel's,
              // even when they carry no values.
              const labelWeight = isTotalRow || displayRow.kind === "header" || displayRow.kind === "collapsed" ? "font-bold text-ink" : "font-medium text-ink";
              // Softer zebra tint + a hover state (2026-10-12, "شكله قذر") — a plain alternating
              // gray with no hover feedback read as flat/dead; the hover highlight is purely
              // cosmetic (selectedCell's own ring is still the real "this is the active cell"
              // signal) but makes the grid feel alive under the mouse the way a real spreadsheet does.
              // Three distinct tiers (2026-10-20 UI refresh): body rows white/zebra; a subtotal is a
              // slate-100 band; the Grand Total is the one gold-accented row (soft gold fill, a 2px
              // gold rule above it) so it reads as the table's bottom line.
              const rowTint =
                displayRow.kind === "grand"
                  ? "bg-gold-soft hover:bg-gold-soft/80 border-t-2 border-t-gold"
                  : isTotalRow
                    ? "bg-slate-100 hover:bg-slate-200/70"
                    : `${displayIndex % 2 === 1 ? "bg-slate-50/60" : "bg-white"} hover:bg-gold-soft/50`;

              // Excel's +/- sits right next to an item's own label, for every field but the
              // innermost (2026-10-15 — it used to be the outermost field only): on its item row in
              // compact/outline form, on its first row in tabular form, and on the collapsed row.
              const collapseToggleFor = (line) => {
                if (!showExpandButtons) return null;
                if (displayRow.kind === "subtotal" || displayRow.kind === "grand" || displayRow.path.length <= line.level) return null;
                const key = collapseKey(line.level, displayRow.path);
                if (!collapsible.has(key)) return null;
                return (
                  <button
                    type="button"
                    onClick={() => toggleGroup(key)}
                    className="text-muted hover:text-ink"
                    aria-label={collapsed.has(key) ? "Expand group" : "Collapse group"}
                  >
                    {collapsed.has(key) ? <ChevronRight className="h-3.5 w-3.5" /> : <ChevronDown className="h-3.5 w-3.5" />}
                  </button>
                );
              };

              return (
                <Fragment key={displayIndex}>
                <tr className={`${rowTint} transition-colors`}>
                  <td className={rowIndex !== null && selectedRows.includes(rowIndex) ? `${cellBorder} bg-slate-200 px-2 py-1.5 text-center text-xs font-semibold text-ink` : rowNumberBodyClass}>
                    {headerRowCount + displayIndex + 1}
                  </td>
                  {rowHeaderGroups.map((group, groupIndex) => {
                    const lines = displayRow.groupCells[groupIndex]?.lines ?? [];
                    // Same degenerate-depth-1 fix as the title row above — a drilldown hierarchy
                    // always goes through the merged branch below (where the +/- toggle lives), even
                    // with only 1 level today, since a SIBLING row could still be deeper once
                    // something else gets expanded.
                    // A combined drilldown slot's own column (Drilldown v3) goes through the line
                    // renderer below too — it's where the depth indent and the +/- toggle live.
                    if (group.start === group.end && !drilldownAlone && group.start !== combinedDrillSlot) {
                      const labelSpan = labelSpans?.[displayIndex][groupIndex] ?? -1;
                      if (labelSpan === 0) return null; // covered by the merged label above
                      return (
                        // Left-aligned like Excel's own row labels in every layout (was centered) — an
                        // outer item's label is merged and centered both ways with "Merge and center
                        // cells with labels" (labelSpan >= 1).
                        <td
                          key={group.start}
                          rowSpan={labelSpan > 1 ? labelSpan : undefined}
                          className={`${cellBorder} whitespace-nowrap px-3 py-1.5 ${labelSpan >= 1 ? "bg-white text-center align-middle" : "text-left"} ${labelWeight}`}
                          {...(lines[0] ? labelHandlers(lines[0].level, displayRow) : {})}
                        >
                          <span className="inline-flex items-center gap-1.5">
                            {lines[0] && collapseToggleFor(lines[0])}
                            {lines[0]?.text ?? ""}
                          </span>
                        </td>
                      );
                    }

                    // A merged (compact-style) column spanning 2+ fields — one indented line per
                    // label buildDisplayRows decided to show here.
                    return (
                      <td key={group.start} className={`${cellBorder} whitespace-nowrap px-3 py-1.5 text-left ${labelWeight}`}>
                        <span className="inline-flex items-center gap-1.5">
                          <span>
                            {lines.map((line) => {
                              const level = line.level;
                              const label = line.text;
                              // Real Excel's per-member OLAP drilldown (PLAN 8, EXPERIMENTAL) — the
                              // +/- toggle sits on the row's own DEEPEST label (its actual current
                              // member, whatever level that happens to be for this particular row —
                              // ragged depth means that's not always the same level across rows),
                              // never on an ancestor label that happens to render in the same merged
                              // column. uniqueNames walks the same ancestry as labels (see
                              // CubeQueryService.BuildDrilldownHeaders), so the last entry is always
                              // this row's own member.
                              //
                              // Drilldown v3: combined with other Rows fields, the member is the one
                              // at the slot (combinedDrillSlot), and every row has one label per field.
                              const memberLevel = drilldownAlone ? labelRow?.labels.length - 1 : combinedDrillSlot;
                              const isOwnMember =
                                rowsHierarchyDrilldown && labelRow && !labelRow.isTotal && displayRow.kind !== "subtotal" && level === memberLevel;
                              const rawUniqueName = isOwnMember ? labelRow.uniqueNames?.[memberLevel] : null;
                              // Real data-quality issue caught live, 2026-10-13: some members on this
                              // cube have an incomplete UniqueName (e.g. ending in a bare "&" with no
                              // key after it — a malformed/empty key on the cube side, not a bug in
                              // this app's own ancestry walk). Every genuine MDX unique-name segment
                              // ends in "]" (either "[x]" or "&[x]") — a value that doesn't is
                              // rejected here instead of being sent to the backend, where it would
                              // fail validation and break the WHOLE query for every row, not just
                              // this one member's own toggle.
                              const uniqueName = rawUniqueName && rawUniqueName.endsWith("]") ? rawUniqueName : null;
                              // Per outer item when combined (Drilldown v3) — the caller knows the
                              // expansion state; the plain list is the drilldown-alone fallback.
                              const expanded = uniqueName
                                ? (isDrilldownMemberExpanded?.(uniqueName, labelRow) ?? rowsHierarchyDrilldown.expandedMemberUniqueNames.includes(uniqueName))
                                : false;
                              return (
                                <div key={level} style={{ paddingLeft: line.indent * 16 }} className="flex items-center gap-1" {...labelHandlers(level, displayRow)}>
                                  {uniqueName && onToggleDrilldownMember && showExpandButtons && (
                                    <button
                                      type="button"
                                      onClick={() => onToggleDrilldownMember(uniqueName, labelRow)}
                                      className="text-muted hover:text-ink"
                                      aria-label={expanded ? "Collapse" : "Expand"}
                                    >
                                      {expanded ? (
                                        <ChevronDown className="h-3 w-3" />
                                      ) : (
                                        <ChevronRight className="h-3 w-3" />
                                      )}
                                    </button>
                                  )}
                                  {/* An outer item's own collapse +/- (none on a drilldown member, which has its own). */}
                                  {!uniqueName && collapseToggleFor(line)}
                                  {label}
                                </div>
                              );
                            })}
                          </span>
                        </span>
                      </td>
                    );
                  })}
                  {rowIndex === null && columnHeaders.map((_, colIndex) => <td key={colIndex} className={`${cellBorder} px-3 py-1.5`} />)}
                  {rowIndex !== null && cells[rowIndex].map((cell, colIndex) => {
                    const isTotalCell = row.isTotal || columnHeaders[colIndex].isTotal;
                    // Which PivotValueField a cell belongs to is read directly off whichever axis
                    // actually carries Σ Values, via that header's own caption at its position —
                    // mirrors the backend's PivotModels.MeasureIndexOf exactly (replaces the old
                    // r/c modulo-stride math, only valid when measures were always innermost).
                    const measureIndex = measuresInRows
                      ? measureIndexOf(row, rowsMeasuresIndex, captionToIndex)
                      : measureIndexOf(columnHeaders[colIndex], columnsMeasuresIndex, captionToIndex);
                    const rangeIndex = measuresInRows ? rowIndex : colIndex;
                    const valueField = valueFields?.length && measureIndex !== null ? valueFields[measureIndex] : null;
                    // Same rule as the saved HTML snapshot: Total/Subtotal cells never get a
                    // conditional-format background, since they're not "a value in the range" —
                    // they're the sum of it (see ConditionalFormatting.cs for the full reasoning).
                    const conditionalStyle = !isTotalCell && valueField
                      ? conditionalStyleFor(cell, valueField.conditionalFormat, conditionalRanges[rangeIndex]?.min, conditionalRanges[rangeIndex]?.max)
                      : null;
                    const iconColor = !isTotalCell && valueField
                      ? iconColorFor(cell, valueField.conditionalFormat, conditionalRanges[rangeIndex]?.min, conditionalRanges[rangeIndex]?.max)
                      : null;

                    const isSelected = selectedCell?.row === rowIndex && selectedCell?.col === colIndex;
                    const isInRange = inRange(rowIndex, colIndex);

                    return (
                      <td
                        key={colIndex}
                        className={`${cellBorder} whitespace-nowrap px-3 py-1.5 text-center tabular-nums ${
                          isTotalCell ? "font-bold text-ink" : "text-slate-700"
                        } ${isSelected ? "relative z-20 ring-2 ring-inset ring-gold-strong" : ""} ${isInRange ? "!bg-gold/25" : ""}`}
                        style={conditionalStyle ?? undefined}
                        // Excel's drag-to-select (2026-10-20): press on a cell and drag over others to
                        // select the range — preventDefault keeps the browser from selecting the TEXT.
                        onMouseDown={(e) => {
                          if (e.button !== 0) return;
                          e.preventDefault();
                          selectCell({ row: rowIndex, col: colIndex }, e.shiftKey);
                          draggingRef.current = true;
                          gridRef.current?.focus();
                        }}
                        onMouseOver={() => {
                          if (draggingRef.current) setSelectedCell({ row: rowIndex, col: colIndex });
                        }}
                        // Excel's double-click on a value cell = Show Details (2026-10-19); with "Enable
                        // show details" off, Excel answers with a message instead.
                        onDoubleClick={() => {
                          if (!request || !valueField || !cellHasDetails(request, valueField, row, columnHeaders[colIndex])) return;
                          if (pivotOptions.enableShowDetails === false) {
                            setNotice("Show details is turned off for this PivotTable (PivotTable Options > Data).");
                            return;
                          }
                          runDrillthrough({ row, column: columnHeaders[colIndex], valueField });
                        }}
                        onContextMenu={(e) => {
                          if (!request || !onChange || !valueField) return;
                          e.preventDefault();
                          // A right-click inside the selected range keeps it (its Copy copies the range).
                          if (!isInRange) selectCell({ row: rowIndex, col: colIndex });
                          setCellMenu({
                            x: e.clientX,
                            y: e.clientY,
                            row,
                            column: columnHeaders[colIndex],
                            valueField,
                            measureIndex,
                            // Raw value only — NOT the formatted string (see this menu's own render
                            // site for why: Copy needs the CURRENT format, re-applied live, the same
                            // way valueField itself is already re-read live instead of frozen at
                            // right-click time).
                            rawValue: cell,
                          });
                        }}
                      >
                        {iconColor && <span style={{ color: iconColor }}>●</span>} {cellText(cell, formatOfCell(rowIndex, colIndex, measureIndex), pivotOptions)}
                      </td>
                    );
                  })}
                </tr>
                {/* Real Excel's per-field "Insert blank line after each item" — placed by
                    buildDisplayRows after the LAST row of a flagged item (its subtotal included). */}
                {displayRow.blankLineAfter && (
                  <tr>
                    <td colSpan={1 + rowHeaderColumnCount + Math.max(1, columnHeaders.length)} className="h-2 border-0 p-0" />
                  </tr>
                )}
                </Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
      {cellMenu && request && onChange && (
        <ValueCellContextMenu
          x={cellMenu.x}
          y={cellMenu.y}
          onClose={() => setCellMenu(null)}
          request={request}
          onChange={onChange}
          // Read LIVE off the current request.values, not the snapshot captured at right-click time
          // (cellMenu.valueField) — every edit made inside this menu (Number Format, Show Values As,
          // Conditional Formatting) already correctly updates the real request, but the menu's OWN
          // inputs were still displaying the STALE right-click-time value back at themselves, so
          // typing a new Decimal Places value (for example) visibly "stuck" at whatever it was when
          // the menu opened — caught live, 2026-10-06 ("لما بكتب حاجة بيفضل الصفر باين"). Falls back
          // to the snapshot only if the index somehow no longer resolves (shouldn't happen — removing
          // this value closes the menu via its own onClose first).
          valueField={request.values[cellMenu.measureIndex] ?? cellMenu.valueField}
          measureIndex={cellMenu.measureIndex}
          sortAxis={measuresInRows ? "Columns" : "Rows"}
          row={cellMenu.row}
          column={cellMenu.column}
          onShowDetails={() => runDrillthrough(cellMenu)}
          onRefresh={() => onRefresh?.()}
          // Re-formatted live off the CURRENT request.values, same reasoning as valueField just
          // above — a frozen string captured at right-click time used to make "Copy" paste the
          // pre-edit value (e.g. "120,000.00") after changing Decimal Places inside this still-open
          // menu to something that should display as "120,000" (caught 2026-10-06, full codebase
          // review).
          displayValue={cellText(cellMenu.rawValue, (request.values[cellMenu.measureIndex] ?? cellMenu.valueField)?.format, pivotOptions)}
          onHideFieldList={() => onHideFieldList?.()}
          onOpenPivotTableOptions={onOpenPivotTableOptions}
          showDetailsEnabled={pivotOptions.enableShowDetails !== false}
          onCopyRange={isMultiCellRange ? copyRange : null}
        />
      )}
      {labelMenu && request && onChange && (
        <RowLabelContextMenu
          {...labelMenu}
          request={request}
          onChange={onChange}
          onClose={() => setLabelMenu(null)}
          fieldCollapse={showExpandButtons ? fieldCollapse : null}
          onToggleItem={labelMenu.collapseKey && collapsible.has(labelMenu.collapseKey) ? () => toggleGroup(labelMenu.collapseKey) : null}
          itemCollapsed={labelMenu.collapseKey ? collapsed.has(labelMenu.collapseKey) : false}
          onMove={(where) => moveItem(labelMenu.level, labelMenu.path, where)}
          onRefresh={onRefresh ? () => onRefresh() : null}
          onOpenPivotTableOptions={onOpenPivotTableOptions}
          onHideFieldList={onHideFieldList ? () => onHideFieldList() : null}
        />
      )}
      {drillthrough && (
        <Modal title={`Show Details — ${drillthrough.label}`} onClose={() => setDrillthrough(null)}>
          {drillthrough.loading ? (
            <p className="text-sm text-muted">Loading…</p>
          ) : drillthrough.error ? (
            <p className="text-sm text-red-600">{drillthrough.error}</p>
          ) : drillthrough.data.rows.length === 0 ? (
            <p className="text-sm text-muted">No underlying rows found for this cell.</p>
          ) : (
            <>
              {/* Real Excel's own warning when Show Details hits its row cap ("Only the first 1000
                  records...") — MAXROWS 1000 (see MdxPivotQueryBuilder.BuildDrillthrough) caps the
                  RESULT at exactly 1000 rows with no way to also ask the cube for the true total
                  underneath the cell, so "got exactly the cap" is the same heuristic real Excel
                  itself uses to decide whether to show this warning. */}
              {drillthrough.data.truncated && (
                <p className="mb-2 rounded-lg border border-amber-300 bg-amber-50 px-3 py-2 text-xs text-amber-800">
                  Showing only the first {drillthrough.data.rows.length.toLocaleString()} rows — there may be more underlying this cell.
                </p>
              )}
              <div className="mb-3 flex items-center gap-3">
                <button
                  type="button"
                  onClick={exportDrillthrough}
                  disabled={drillthroughExporting}
                  className="btn-secondary inline-flex items-center gap-1.5 disabled:cursor-not-allowed disabled:opacity-50"
                >
                  <FileSpreadsheet className="h-4 w-4" style={{ color: "#217346" }} />
                  {drillthroughExporting ? "Exporting…" : "Export to Excel"}
                </button>
                <span className="text-xs text-muted">{drillthrough.data.rows.length.toLocaleString()} rows</span>
                {drillthrough.exportError && <span className="text-xs font-medium text-red-700">{drillthrough.exportError}</span>}
              </div>
              <table className="border-collapse text-sm">
              <thead>
                <tr>
                  {drillthrough.data.columns.map((col, i) => (
                    <th key={i} className="border border-slate-300 bg-slate-100 px-3 py-1.5 text-left font-semibold text-ink">
                      {col}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {drillthrough.data.rows.map((dataRow, rowIdx) => (
                  <tr key={rowIdx} className={rowIdx % 2 === 0 ? "bg-slate-50/60" : ""}>
                    {dataRow.map((value, colIdx) => (
                      <td key={colIdx} className="border border-slate-300 px-3 py-1.5 text-slate-700">
                        {value ?? ""}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
              </table>
            </>
          )}
        </Modal>
      )}
    </>
  );
}
