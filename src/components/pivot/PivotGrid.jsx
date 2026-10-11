import { useEffect, useMemo, useRef, useState } from "react";
import { ChevronDown, ChevronRight, ChevronLeft, Filter as FilterIcon, Table2, X } from "lucide-react";
import Popover from "../forms/Popover";
import ContextMenu from "../forms/ContextMenu";
import Modal from "../forms/Modal";
import Slicer from "./Slicer";
import Timeline from "./Timeline";
import { SortEditor, FilterRow, NumberFormatEditor, ValueFieldSettingsPanel, ConditionalFormatEditor, SHOW_VALUES_AS_OPTIONS, PERCENT_SHOW_VALUES_AS, DEFAULT_FORMAT } from "./FieldPicker";
import { formatNumber, formatForColumn } from "../../utils/numberFormat";
import { formatDimensionName } from "../../utils/cubeMeta";
import { conditionalStyleFor, detailRangeForColumn, detailRangeForRow, iconColorFor } from "../../utils/conditionalFormat";
import { computeGroups, computeCompactVisibility } from "../../utils/pivotGridLayout";
import { MEASURES_PLACEHOLDER, measuresIndexOf, buildCaptionToIndex, measureIndexOf } from "../../utils/measuresPlacement";
import { drillthroughPivot } from "../../services/pivotApi";

// Excel's own Report Filter area: every field in Filters gets its own row ABOVE the PivotTable
// itself — field name on the left, current selection (plus a dropdown to change it) on the right
// — confirmed live 2026-10-05 against a real Excel screenshot. This was previously only editable
// from the separate Filters box in the field-picker sidebar; Excel shows (and lets you re-filter
// from) it right here too, which this app didn't have at all.
//
// "(All)" when nothing is actually narrowing the field yet (e.g. Members mode with no members
// picked), the single member's own name when exactly one is picked, "(Multiple Items)" for more
// than one or for any of the other filter modes (label/TopN/range) — same convention real Excel
// itself uses for a Report Filter whose selection isn't a single plain value.
function filterValueSummary(filter) {
  if (filter.mode === "Members" || filter.mode === "Exclude") {
    if (filter.includedMembers.length === 0) return "(All)";
    if (filter.includedMembers.length === 1) return filter.includedMembers[0];
    return "(Multiple Items)";
  }
  if (filter.mode === "LabelContains" || filter.mode === "LabelBeginsWith" || filter.mode === "LabelEndsWith") {
    return filter.labelText ? "(Multiple Items)" : "(All)";
  }
  return "(Multiple Items)";
}

function FilterHeaderRow({ filter, index, request, onChange, dimensionOptions, selectedValueOptions }) {
  const updateFilter = (patch) => {
    const next = [...request.filters];
    next[index] = { ...next[index], ...patch };
    onChange({ ...request, filters: next });
  };
  const removeFilter = () => {
    onChange({ ...request, filters: request.filters.filter((_, i) => i !== index) });
  };

  return (
    <div className="flex items-stretch border-b border-slate-100 last:border-b-0">
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
function HeaderCell({ children, muted = false, isTotal = false }) {
  return (
    <span
      className={`flex items-center justify-center gap-1 text-sm font-semibold ${
        muted ? "text-muted" : isTotal ? "text-gold-strong" : "text-ink"
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
function FieldHeaderMenu({ label, field, axis, request, onChange, muted = false, isTotal = false, level = 0 }) {
  // "More Sort Options…"/"Field Settings…" open the SAME panel in place of the plain shortcut
  // list, same back-and-forth pattern as Excel's own header dropdown (its "More Sort Options..."
  // item opens a bigger dialog from the same arrow) — relocated here from FieldPicker's toolbar
  // entirely, since Excel never puts per-field Sort/"Show items with no data" in a side task pane,
  // only on the field's own dropdown arrow or its right-click menu (moved 2026-10-04, "المفروض
  // تبقى في نفس مكان الـ excel").
  const [panel, setPanel] = useState("shortcuts");

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
  const noDataKey = axis === "Columns" ? "Columns" : "Rows";
  const showsNoDataItems = request.showItemsWithNoData.includes(noDataKey);

  const setSort = (direction) => {
    // `sort` is mutually exclusive with RowLevelSorts/ColumnLevelSorts GLOBALLY, not just on this
    // axis (MdxPivotQueryBuilder.ValidateLevelSorts rejects `sort` whenever EITHER array is
    // non-empty) — both PivotBuilder's own per-axis default-sort-on-first-field and the
    // LevelSortEditor can leave the OTHER axis's LevelSorts populated even while sorting on this
    // one, so both must be cleared here, not just this axis's (caught live, 2026-10-06: a plain
    // two-axis pivot defaults BOTH axes' LevelSorts, so the very first click of this header's own
    // Sort shortcut on either axis was throwing "Could not query the cube" otherwise).
    onChange({ ...request, sort: { direction, byMeasureField: null, axis, customOrder: null }, rowLevelSorts: [], columnLevelSorts: [] });
  };

  const clearSort = () => {
    if (isLevelSortedHere) {
      const key = axis === "Columns" ? "columnLevelSorts" : "rowLevelSorts";
      onChange({ ...request, [key]: levelSorts.filter((s) => s.level !== level) });
      return;
    }
    onChange({ ...request, sort: null });
  };

  const addFilterShortcut = () => onChange({ ...request, filters: [...request.filters, { field, mode: "Members", includedMembers: [] }] });

  const removeFilter = () => onChange({ ...request, filters: request.filters.filter((f) => f.field !== field) });

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
      panelClassName="min-w-[14rem] max-w-[18rem]"
      trigger={(toggle) => (
        <button
          type="button"
          onClick={() => {
            setPanel("shortcuts");
            toggle();
          }}
          className={`flex w-full items-center justify-center gap-1 text-sm font-semibold hover:text-gold ${
            muted ? "text-muted" : isTotal ? "text-gold-strong" : "text-ink"
          }`}
          title={`Sort or filter ${label}`}
        >
          {label}
          {hasFilter && <FilterIcon className="h-3 w-3 shrink-0 text-gold" />}
          <ChevronDown className="h-3.5 w-3.5 shrink-0" />
        </button>
      )}
    >
      {(close) =>
        panel === "sort" ? (
          <div>
            <button
              type="button"
              onClick={() => setPanel("shortcuts")}
              className="mb-2 inline-flex items-center gap-1 text-xs font-medium text-muted hover:text-ink"
            >
              <ChevronLeft className="h-3.5 w-3.5" /> Back
            </button>
            <SortEditor
              sort={request.sort}
              rowLevelSorts={request.rowLevelSorts}
              columnLevelSorts={request.columnLevelSorts}
              values={request.values}
              rows={request.rows}
              columns={request.columns}
              onChange={(patch) => onChange({ ...request, ...patch })}
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
            <button type="button" onClick={() => setPanel("sort")} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
              More Sort Options…
            </button>
            <div className="my-1 border-t border-slate-100" />
            {hasFilter ? (
              <button
                type="button"
                onClick={() => {
                  removeFilter();
                  close();
                }}
                className="rounded-lg px-2 py-1.5 text-left text-muted hover:bg-slate-100"
              >
                Remove filter
              </button>
            ) : (
              <button
                type="button"
                onClick={() => {
                  addFilterShortcut();
                  close();
                }}
                className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10"
              >
                Add filter…
              </button>
            )}
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
          </div>
        )
      }
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
function ValueCellContextMenu({ x, y, onClose, request, onChange, valueField, measureIndex, sortAxis, row, column, onShowDetails, onRefresh, displayValue, onHideFieldList }) {
  const [panel, setPanel] = useState("root");
  const format = valueField.format ?? DEFAULT_FORMAT;
  const isDimensionField = valueField.field.startsWith("[");
  const displayLabel = valueField.customLabel || (isDimensionField ? formatDimensionName(valueField.field) : valueField.field);

  // `options.requireGrandTotals` rides along in the SAME onChange call as the value patch — see
  // ValueFieldSettingsPanel's own Show Values As handler for why this can't be a separate call.
  // Forces the safe universal superset (RowsAndColumns) rather than computing which SPECIFIC axis
  // this particular ShowValuesAs variant actually needs (PercentOfGrandTotal/PercentOfParentRow:
  // Rows; PercentOfParentColumn: Columns; Index: both) — simpler, and always satisfies whichever one
  // applies, at the cost of also turning on the OTHER axis's total even when not strictly required.
  const updateValue = (patch, options = {}) => {
    const values = request.values.map((v, i) => (i === measureIndex ? { ...v, ...patch } : v));
    onChange({ ...request, values, ...(options.requireGrandTotals ? { grandTotals: "RowsAndColumns" } : {}) });
  };
  const updateFormat = (patch) => updateValue({ format: { ...format, ...patch } });
  // Same global-exclusivity clear as FieldHeaderMenu's own setSort above — this menu's Sort
  // shortcut hits the identical "both axes default-sorted, this throws otherwise" case.
  const setSort = (direction) =>
    onChange({ ...request, sort: { direction, byMeasureField: valueField.field, axis: sortAxis, customOrder: null }, rowLevelSorts: [], columnLevelSorts: [] });
  const removeValue = () => onChange({ ...request, values: request.values.filter((_, i) => i !== measureIndex) });

  const isCalculatedField = request.calculatedFields.some((cf) => cf.name === valueField.field);
  // A Calculated Item is a synthetic MDX member (e.g. "East + West" combined into one row/column)
  // with a real, non-null UniqueName of its own — NOT backed by a single real fact-table row set
  // the same way a Calculated Field isn't, so it needs the exact same exclusion here (caught
  // 2026-10-06, full codebase review: Show Details was still offered on a Calculated Item cell,
  // whose drillthrough would at best return only one of the combined members' own rows).
  const isCalculatedItem = Boolean(row.isCalculatedItem) || Boolean(column.isCalculatedItem);
  const canShowDetails = Boolean(row.uniqueNames) && Boolean(column.uniqueNames) && valueField.showValuesAs === "Normal" && !isDimensionField && !isCalculatedField && !isCalculatedItem;

  return (
    <ContextMenu x={x} y={y} onClose={onClose}>
      {panel === "root" ? (
        <div className="flex flex-col gap-1">
          <button
            type="button"
            onClick={async () => {
              onClose();
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
              onClick={() => { onClose(); onShowDetails(); }}
              className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10"
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
          <button type="button" onClick={() => setPanel("pivotTableOptions")} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
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
        <div>
          <button
            type="button"
            onClick={() => setPanel("sort")}
            className="mb-2 inline-flex items-center gap-1 px-2 text-xs font-medium text-muted hover:text-ink"
          >
            <ChevronLeft className="h-3.5 w-3.5" /> Back
          </button>
          {/* Same full SortEditor the field-header dropdown's own "More Sort Options…" opens
              (FieldHeaderMenu above) — this cell menu's Sort flyout only had the two direct
              Smallest/Largest shortcuts until now, unlike the field-header one, which is an
              inconsistency real Excel doesn't have (its own cell Sort submenu has this same item
              too) — caught during a full audit pass, 2026-10-05. */}
          <SortEditor
            sort={request.sort}
            rowLevelSorts={request.rowLevelSorts}
            columnLevelSorts={request.columnLevelSorts}
            values={request.values}
            rows={request.rows}
            columns={request.columns}
            onChange={(patch) => onChange({ ...request, ...patch })}
          />
        </div>
      ) : panel === "numberFormat" ? (
        <NumberFormatEditor format={format} onUpdateFormat={updateFormat} onBack={() => setPanel("root")} />
      ) : panel === "settings" ? (
        <ValueFieldSettingsPanel v={valueField} onUpdate={updateValue} onUpdateFormat={updateFormat} onBack={() => setPanel("root")} rows={request.rows} />
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
              // dropdown (caught 2026-10-06, full codebase review).
              disabled={request.rows.filter((f) => f !== MEASURES_PLACEHOLDER).length < 2}
              onChange={(e) => onChange({ ...request, subtotals: e.target.value })}
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
                if (PERCENT_SHOW_VALUES_AS.has(opt.value) && format.type === "General") {
                  valuePatch.format = { ...format, type: "Percentage" };
                }
                const values = request.values.map((v, i) => (i === measureIndex ? { ...v, ...valuePatch } : v));
                // One single onChange call for both the value patch AND grandTotals — two
                // separate calls (even via updateValue + a second onChange) would each build off
                // the same stale `request` closure, so whichever ran second would silently discard
                // the other (same race this file has already hit and fixed elsewhere).
                const requiresGrandTotals = opt.value === "PercentOfGrandTotal" || opt.value === "Index";
                onChange({ ...request, values, ...(requiresGrandTotals ? { grandTotals: "RowsAndColumns" } : {}) });
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
}) {
  const [collapsed, setCollapsed] = useState(() => new Set());
  // The right-clicked value cell's context menu ({x, y, row, column, valueField, measureIndex}) and
  // the Show Details modal's own state (fetched separately, since the menu closes immediately on
  // click but the modal needs to stay open through the fetch and after) — see ValueCellContextMenu
  // above and the contextmenu handler / modal render below.
  const [cellMenu, setCellMenu] = useState(null);
  const [drillthrough, setDrillthrough] = useState(null);
  // Keyboard navigation (Excel's own arrow-key cell selection, Enter to move down, Escape to
  // clear) — {row, col} are indices into `cells`/`rowHeaders`/`columnHeaders`, not screen position,
  // so the selection survives sorting/filtering as long as the same indices still exist. Only the
  // VALUE cells are navigable (not headers) — this app's grid isn't an editable spreadsheet, so the
  // one thing worth replicating is "move the active cell around with arrow keys to read values
  // without reaching for the mouse," not full worksheet-wide navigation.
  const [selectedCell, setSelectedCell] = useState(null);
  // A fresh `result` (re-sort, re-filter, a field added/removed) can reshuffle which row/column
  // lands at a given index — `selectedCell` is plain {row, col} ARRAY INDICES with no identity
  // check against what's actually there now, so the ring highlight used to silently keep "selecting"
  // whatever row/column now happens to sit at the SAME indices in the new, unrelated result (caught
  // 2026-10-06, full codebase review). Cleared on every new result instead of carried forward.
  useEffect(() => {
    setSelectedCell(null);
  }, [result]);
  const gridRef = useRef(null);
  // Same shape FieldPicker.jsx builds for FilterRow's own "choose a field"/"by measure" selects —
  // see FilterHeaderRow above, which reuses FilterRow unchanged for the Report Filter area's own
  // dropdown.
  const dimensionOptions = dimensions.map((d) => ({ value: d.field, label: d.displayName }));
  const selectedValueOptions = valueFields.filter((v) => v.field).map((v) => ({ value: v.field, label: v.field }));

  const groups = useMemo(() => (result ? computeGroups(result.rowHeaders) : []), [result]);
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
  // Outline AND Tabular reuse the exact same visibility computation as Compact (see
  // computeCompactVisibility's own comment) — they just render the result into separate <td> cells
  // per level instead of one merged, indented cell. Real Excel's own default for Outline/Tabular is
  // to suppress a repeated outer-level label (same as Compact) — "Repeat All Item Labels" is a
  // separate, explicit opt-in that forces every level to show on every row instead (confirmed live,
  // 2026-10-07, against a real Excel screenshot). Compact never responds to this toggle — real
  // Excel grays the option out entirely for Compact Form, since there's no separate per-level
  // column to repeat a label into.
  const suppressedLevelVisibility = useMemo(
    () =>
      result && (layout === "Compact" || ((layout === "Outline" || layout === "Tabular") && !repeatItemLabels))
        ? computeCompactVisibility(result.rowHeaders)
        : null,
    [result, layout, repeatItemLabels],
  );

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
        filters: request.filters,
      });
      setDrillthrough({ loading: false, error: null, data, label });
    } catch (err) {
      setDrillthrough({ loading: false, error: err.message, data: null, label });
    }
  };

  // Excel never leaves the worksheet blank before the first field is placed — it shows this exact
  // placeholder inside the PivotTable's own bounds ("To build a report, choose fields from the
  // PivotTable Field List"). Matching that instead of rendering nothing (flagged live, 2026-10-04,
  // against a side-by-side screenshot of a brand-new real Excel PivotTable).
  if (!result) {
    return (
      <div className="flex min-h-[22rem] items-center justify-center rounded-xl border-2 border-dashed border-slate-300 bg-white/60">
        <div className="max-w-xs text-center">
          <Table2 className="mx-auto mb-3 h-10 w-10 text-gold/50" />
          <p className="text-sm font-semibold text-ink">To build a report, choose fields from the Fields panel</p>
          <p className="mt-1 text-xs text-muted">Drag a field into Filters, Columns, Rows or Values to get started.</p>
        </div>
      </div>
    );
  }

  const { rowHeaders, columnHeaders, cells } = result;
  const rowDepth = rowHeaders[0]?.labels.length ?? 0;
  const rowHeaderColumnCount = layout === "Compact" ? 1 : rowDepth;
  // Backend floor of 1 (not 0) whenever there's at least one row to label: with no Value fields at
  // all (Excel's own "removed the last Value field" state), Columns can legitimately come back with
  // ZERO headers (a literal empty MDX set — see MdxPivotQueryBuilder.BuildAxisFinal's own WHY
  // comment) — but the level-0 header row below is also where the ROW field's own title cell
  // (FieldHeaderMenu "Business Unit.BU Short Name" etc.) renders, via the `Array.from({length:
  // columnDepth})` loop a few lines down. columnDepth === 0 would skip that loop entirely and drop
  // the row headers' own title row along with the (correctly absent) column ones — caught live,
  // 2026-10-05, the first time this zero-Values case became reachable at all.
  const columnDepth = columnHeaders[0]?.labels.length ?? (rowHeaders.length > 0 ? 1 : 0);

  const toggleGroup = (key) => {
    setCollapsed((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  const groupForRow = new Map();
  groups.forEach((g) => {
    for (let r = g.start; r < g.end; r++) groupForRow.set(r, g);
  });

  const isRowVisible = (row, rowIndex) => {
    const group = groupForRow.get(rowIndex);
    const isGroupSubtotalRow = group?.hasSubtotal && rowIndex === group.end - 1;
    const isCollapsedDetail = group?.hasSubtotal && collapsed.has(group.key) && !isGroupSubtotalRow;
    return !isCollapsedDetail;
  };

  // Precomputed rather than mutated inline during the render map below, so the zebra-striping
  // position stays a pure function of the current collapse state instead of a running counter touched
  // mid-render.
  const visibleRowPosition = new Map();
  {
    let counter = 0;
    rowHeaders.forEach((row, rowIndex) => {
      if (isRowVisible(row, rowIndex)) {
        counter += 1;
        visibleRowPosition.set(rowIndex, counter);
      }
    });
  }

  // Grid lines on every header/body cell (border on all four sides, not just a per-row bottom
  // border) — the single biggest visual gap vs a real spreadsheet grid (flagged live, 2026-09-28).
  const cellBorder = "border border-slate-300";

  // Excel's own arrow-key cell navigation — see selectedCell's own WHY comment above. Skips over
  // collapsed/hidden rows the same way the mouse already does (clicking a collapsed group's hidden
  // detail rows isn't possible either, since they're not rendered at all).
  const handleGridKeyDown = (e) => {
    if (!["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight", "Enter", "Escape"].includes(e.key)) return;
    e.preventDefault();
    if (e.key === "Escape") {
      setSelectedCell(null);
      return;
    }
    if (!selectedCell) {
      const firstVisibleRow = rowHeaders.findIndex((row, i) => isRowVisible(row, i));
      if (firstVisibleRow >= 0 && columnHeaders.length > 0) setSelectedCell({ row: firstVisibleRow, col: 0 });
      return;
    }
    const { row, col } = selectedCell;
    if (e.key === "ArrowLeft") {
      setSelectedCell({ row, col: Math.max(0, col - 1) });
    } else if (e.key === "ArrowRight") {
      setSelectedCell({ row, col: Math.min(columnHeaders.length - 1, col + 1) });
    } else if (e.key === "ArrowUp") {
      for (let r = row - 1; r >= 0; r--) {
        if (isRowVisible(rowHeaders[r], r)) {
          setSelectedCell({ row: r, col });
          return;
        }
      }
    } else if (e.key === "ArrowDown" || e.key === "Enter") {
      for (let r = row + 1; r < rowHeaders.length; r++) {
        if (isRowVisible(rowHeaders[r], r)) {
          setSelectedCell({ row: r, col });
          return;
        }
      }
    }
  };

  return (
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
      {request && onChange && request.filters.some((f) => !request.slicerFields.includes(f.field) && !request.timelineFields.includes(f.field)) && (
        <div className="card mb-2 w-fit max-w-full overflow-hidden p-0">
          {request.filters.map((filter, index) =>
            request.slicerFields.includes(filter.field) || request.timelineFields.includes(filter.field) ? null : (
              <FilterHeaderRow
                key={index}
                filter={filter}
                index={index}
                request={request}
                onChange={onChange}
                dimensionOptions={dimensionOptions}
                selectedValueOptions={selectedValueOptions}
              />
            ),
          )}
        </div>
      )}
      {/* w-fit: the white card now hugs the table's actual (content-sized) width instead of
          stretching to the full page — dropping the table's own w-full (previous fix) stopped
          individual columns from being stretched, but left the surrounding card still full-width
          with a large empty gap on the right, which is what "not sized right" meant here.
          No overflow-hidden anymore either — ANY ancestor with overflow other than visible breaks
          position:sticky for everything inside it (the thead's own sticky top-0 right below, AND
          PivotBuilder.jsx's field-list <aside>, which sits as a SIBLING further up the tree but hits
          the exact same rule) — a non-scrolling overflow:hidden box still counts as "the nearest
          scroll container" for sticky purposes, so a descendant sticky element ends up pinned
          relative to a box that itself never scrolls, which looks identical to sticky just not
          working at all (caught live, 2026-10-06, same general class of bug as Popover.jsx's own
          portal fix earlier this session, just hitting sticky instead of an absolutely-positioned
          panel). The table's own rounded corners (clipped by this overflow-hidden before) are a
          purely cosmetic loss next to a frozen header that doesn't freeze. */}
      <div className="card w-fit max-w-full p-0">
      {/* Back to a bounded, scrollable grid area (max-h + overflow-auto on BOTH axes) — the
          page-scroll version tried just above (removing this box entirely) traded away horizontal
          containment to get it: with no overflow anywhere in the chain, a WIDE result (many date
          columns, say) had nothing to clip or scroll it, so it visually bled out under the sticky
          field-list sidebar instead of being reachable by any scrollbar at all (caught live,
          2026-10-06, after a direct side-by-side with real Excel's own worksheet pane — "مش
          المفروض يبقى في جزء ليه scroll bar زي ال excel كدا" — real Excel's grid IS a bounded,
          independently-scrolling region next to its docked Field List, never a page that grows
          wider than the window). Sized generously (most of the viewport, not the small, cramped box
          from an EARLIER complaint this exact area already went through) rather than tiny — this is
          the balance of both: contained like Excel, not cramped like before. Putting overflow-auto
          on THIS inner div instead of the outer decorative `card` above also means the thead's own
          sticky top-0 below sticks correctly within a real, working scroll container (the standard,
          reliable pattern), rather than relying on the page itself scrolling — no CSS-spec overflow-
          visible/auto-axis-matching quirk to fight either, since this div's height actually IS
          constrained now. tabIndex + onKeyDown here (not on individual cells) is what makes
          arrow-key navigation work — a plain <td> isn't focusable, so a cell click also calls
          gridRef.current.focus() below to move real DOM focus onto this container. */}
      <div
        ref={gridRef}
        className={boundedHeight ? "max-h-[80vh] overflow-auto outline-none" : "outline-none"}
        tabIndex={0}
        onKeyDown={handleGridKeyDown}
      >
        <table className="border-collapse text-sm">
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
          <thead className={boundedHeight ? "sticky top-0 z-10" : undefined}>
            {columnFieldLabels.length > 0 && (
              <tr>
                <th colSpan={rowHeaderColumnCount} className="bg-white" />
                {/* Math.max(1, ...) — columnHeaders can legitimately be EMPTY (a Columns field
                    configured but the last Value field removed, see columnDepth's own WHY comment
                    above) while a Columns field is still configured (columnFieldLabels.length > 0
                    still true) — colSpan={0} isn't a meaningful table attribute and collapsed this
                    cell (with its own working FieldHeaderMenu dropdown) unpredictably instead of
                    just rendering at its practical minimum width (caught 2026-10-06, full codebase
                    review). */}
                <th colSpan={Math.max(1, columnHeaders.length)} className={`${cellBorder} bg-slate-100 px-2 py-2`}>
                  <FieldHeaderMenu
                    label={columnFieldLabels.join(" › ")}
                    field={columns.length === 1 ? columns[0] : null}
                    axis="Columns"
                    request={request}
                    onChange={onChange}
                    muted
                  />
                </th>
              </tr>
            )}
            {Array.from({ length: columnDepth }).map((_, level) => (
              <tr key={level}>
                {level === 0 && (
                  <>
                    {layout === "Compact" ? (
                      // Compact merges every row field into ONE column — not one-to-one with a
                      // single field, so this stays a plain (non-interactive) header, unlike the
                      // per-level headers below. Real Excel's own Compact form always shows the
                      // literal generic caption "Row Labels" here (never the actual field name(s)),
                      // confirmed live 2026-10-05 against a real Excel screenshot — fixed from
                      // joining the field names together, which is what Outline/Tabular correctly
                      // do instead (see the per-level FieldHeaderMenu branch below).
                      <th rowSpan={columnDepth} className={`${cellBorder} bg-slate-100 px-2 py-2`}>
                        <HeaderCell muted>Row Labels</HeaderCell>
                      </th>
                    ) : (
                      Array.from({ length: rowDepth }).map((_, rowLevel) => (
                        <th key={rowLevel} rowSpan={columnDepth} className={`${cellBorder} bg-slate-100 px-2 py-2`}>
                          <FieldHeaderMenu
                            label={rowFieldLabels[rowLevel] ?? ""}
                            field={rows[rowLevel] ?? null}
                            axis="Rows"
                            level={rowLevel}
                            request={request}
                            onChange={onChange}
                            muted
                          />
                        </th>
                      ))
                    )}
                  </>
                )}
                {columnHeaders.map((col, colIndex) => (
                  <th key={colIndex} className={`${cellBorder} px-2 py-2 ${col.isTotal ? "bg-gold/10" : "bg-slate-50"}`}>
                    <HeaderCell isTotal={col.isTotal}>{col.labels[level]}</HeaderCell>
                  </th>
                ))}
              </tr>
            ))}
          </thead>
          <tbody>
            {rowHeaders.map((row, rowIndex) => {
              if (!isRowVisible(row, rowIndex)) {
                return null;
              }

              const group = groupForRow.get(rowIndex);
              const isGroupSubtotalRow = group?.hasSubtotal && rowIndex === group.end - 1;
              const rowTint = row.isTotal ? "bg-gold/10" : visibleRowPosition.get(rowIndex) % 2 === 0 ? "bg-slate-50/60" : "";

              const collapseToggle = isGroupSubtotalRow && (
                <button
                  type="button"
                  onClick={() => toggleGroup(group.key)}
                  className="text-muted hover:text-ink"
                  aria-label={collapsed.has(group.key) ? "Expand group" : "Collapse group"}
                >
                  {collapsed.has(group.key) ? <ChevronRight className="h-3.5 w-3.5" /> : <ChevronDown className="h-3.5 w-3.5" />}
                </button>
              );

              return (
                <tr key={rowIndex} className={rowTint}>
                  {layout === "Compact" ? (
                    <td
                      className={`${cellBorder} whitespace-nowrap px-6 py-2.5 text-left ${row.isTotal ? "font-bold text-ink" : "font-medium text-ink"}`}
                    >
                      <span className="inline-flex items-center gap-1.5">
                        {collapseToggle}
                        <span>
                          {row.labels.map((label, level) =>
                            suppressedLevelVisibility[rowIndex][level] ? (
                              <div key={level} style={{ paddingLeft: level * 16 }}>
                                {label}
                              </div>
                            ) : null,
                          )}
                        </span>
                      </span>
                    </td>
                  ) : (
                    // Outline and Tabular render identically — one <td> per row field — the only
                    // difference is whether a repeated outer-level label is suppressed, which
                    // suppressedLevelVisibility above already resolved to null (show everything,
                    // Repeat All Item Labels is on) or a real per-row/per-level visibility array.
                    row.labels.map((label, level) => (
                      <td
                        key={level}
                        className={`${cellBorder} whitespace-nowrap px-6 py-2.5 text-center ${row.isTotal ? "font-bold text-ink" : "font-medium text-ink"}`}
                      >
                        <span className="inline-flex items-center justify-center gap-1.5">
                          {level === 0 && collapseToggle}
                          {suppressedLevelVisibility ? (suppressedLevelVisibility[rowIndex][level] ? label : "") : label}
                        </span>
                      </td>
                    ))
                  )}
                  {cells[rowIndex].map((cell, colIndex) => {
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

                    return (
                      <td
                        key={colIndex}
                        className={`${cellBorder} whitespace-nowrap px-6 py-2.5 text-center tabular-nums ${
                          isTotalCell ? "font-bold text-ink" : "text-slate-700"
                        } ${isSelected ? "relative z-20 ring-2 ring-inset ring-ink" : ""}`}
                        style={conditionalStyle ?? undefined}
                        onClick={() => {
                          setSelectedCell({ row: rowIndex, col: colIndex });
                          gridRef.current?.focus();
                        }}
                        onContextMenu={(e) => {
                          if (!request || !onChange || !valueField) return;
                          e.preventDefault();
                          setSelectedCell({ row: rowIndex, col: colIndex });
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
                        {iconColor && <span style={{ color: iconColor }}>●</span>} {formatNumber(cell, formatForColumn(measureIndex, valueFields))}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
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
          displayValue={formatNumber(cellMenu.rawValue, (request.values[cellMenu.measureIndex] ?? cellMenu.valueField)?.format)}
          onHideFieldList={() => onHideFieldList?.()}
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
