import { useEffect, useState } from "react";
import Select from "react-select";
import { X, GripVertical, Plus, ChevronDown, Filter, Columns3, Rows3, Sigma, Loader2 } from "lucide-react";
import HierarchicalCubePathSelect from "../forms/HierarchicalCubePathSelect";
import MultiSelectField from "../forms/MultiSelectField";
import Popover from "../forms/Popover";
import { getMembers } from "../../services/cubeMetaApi";
import { formatDimensionName } from "../../utils/cubeMeta";
import { setFieldDragPayload, readFieldDragPayload } from "../../utils/dragDrop";

const SHOW_VALUES_AS_OPTIONS = [
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

const DEFAULT_FORMAT = { type: "General", decimalPlaces: 2, currencySymbol: "EGP" };

const nativeSelectClass =
  "rounded-xl border border-slate-200 bg-white/90 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent";

// A row of colored name+count chips, each toggling whether the full thing it represents is
// shown below — the SAME show/hide-by-picking-a-chip pattern used twice in this file: once for
// which cube TABLES' column lists appear in AvailableFieldsPanel, and once for which ZONE
// SectionCards (Filters/Rows/Columns/Values/...) appear at all (see FieldPicker's own render).
// Extracted here instead of duplicated because the user explicitly asked for the zone picker to
// work "زي الـ cards" (like the [table] cards) after an earlier per-card chevron-collapse
// attempt wasn't what they wanted — one real component, one real behavior, reused twice.
function ToggleChipRow({ items, visibleNames, onToggle, onSelectAll, onClearAll }) {
  return (
    <div className="mb-4 flex flex-wrap items-center gap-2">
      {items.map(({ name, badge }) => {
        const isVisible = visibleNames.includes(name);
        return (
          <button
            key={name}
            type="button"
            onClick={() => onToggle(name)}
            // Solid gold for every chip regardless of table — a distinct rainbow color per table
            // (one of 10 cycling colors) looked inconsistent with the rest of the app's single
            // gold/ink theme, especially once tables became a user-picked, per-browser subset
            // rather than a small fixed handful (flagged live, 2026-09-27).
            className={`flex items-center gap-2.5 rounded-lg bg-gold px-4 py-2.5 text-sm font-bold uppercase tracking-wide text-white shadow-sm transition ${
              isVisible ? "ring-2 ring-ink/25 ring-offset-2" : ""
            }`}
            aria-pressed={isVisible}
            title={isVisible ? `Hide ${name}` : `Show ${name}`}
          >
            {name}
            {badge > 0 && <span className="rounded-full bg-white/30 px-1.5 py-0.5 text-[10px] font-bold leading-none">{badge}</span>}
            <span className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-white/30 text-sm font-bold leading-none">
              {isVisible ? "−" : "+"}
            </span>
          </button>
        );
      })}
      <div className="ml-auto flex shrink-0 gap-2">
        <button type="button" onClick={onSelectAll} className="btn-secondary px-3 py-1.5 text-xs">
          Select all
        </button>
        <button type="button" onClick={onClearAll} className="btn-secondary px-3 py-1.5 text-xs">
          Clear all
        </button>
      </div>
    </div>
  );
}

// Excel's Field List: every cube DIMENSION field in one place, dragged from here into whichever
// of Filters/Rows/Columns it belongs in. Unlike the zones, a field never "leaves" this list when
// placed somewhere — Excel's own field list keeps every field checkbox visible too — it just gets
// a small dot marking it as already used somewhere, since (unlike Excel) this tool allows the same
// field to sit in Filters AND Rows/Columns at once (a Row/Column field can independently carry its
// own member filter — see moveField's WHY comment in FieldPicker).
//
// Measures are deliberately NOT listed here — a measure has exactly one valid destination
// (Values), so dragging added nothing but an extra click and a way to mis-drop a measure into
// Rows/Columns and produce broken MDX (a real bug hit live, 2026-09-23: a measure dragged into
// Values by mistake produced malformed MDX). Picking measures is a plain multi-select in the
// Values box instead — see FieldPicker's own render.
//
// Grouped by dimension (the text before the first "." in displayName, e.g. "Patient" for
// "Patient.Patient Name") rather than one flat wrapped wall of pills — a real cube has 20+ fields
// across 8+ dimensions, and a single ungrouped list of same-weight buttons became unreadable at
// that size (caught live 2026-09-22: "ايه الشكل القذر دا"). Matches Excel's own Field List, which
// is always grouped by table/dimension with a bold header per group, never one flat list.
function AvailableFieldsPanel({ dimensions, isPlaced, tables, tablesLoading, enabledTableNames, onEnabledTableNamesChange, fieldsLoading }) {
  const groups = [];
  const groupIndexByName = new Map();
  const addToGroup = (groupName, field, label) => {
    if (!groupIndexByName.has(groupName)) {
      groupIndexByName.set(groupName, groups.length);
      groups.push({ name: groupName, fields: [] });
    }
    groups[groupIndexByName.get(groupName)].fields.push({ field, label });
  };

  // `dimensions` only ever holds columns for tables PivotBuilder has already fetched (i.e. every
  // currently-enabled table) — so grouping it is enough on its own, no separate filter needed
  // against enabledTableNames here.
  dimensions.forEach((d) => {
    const dotIndex = d.displayName.indexOf(".");
    const groupName = dotIndex === -1 ? d.displayName : d.displayName.slice(0, dotIndex);
    const label = dotIndex === -1 ? d.displayName : d.displayName.slice(dotIndex + 1);
    addToGroup(groupName, d.field, label);
  });

  // No table's COLUMN list shown until the user explicitly picks one (+ chip or "Select all")
  // even among their enabled tables — a handful of tables can still mean 20+ columns each, so
  // opening with every column list expanded is exactly the wall-of-content this redesign exists
  // to avoid.
  const [visibleGroupNames, setVisibleGroupNames] = useState([]);
  const visibleGroups = groups.filter((g) => visibleGroupNames.includes(g.name));

  const toggleGroup = (name) => {
    setVisibleGroupNames((prev) => (prev.includes(name) ? prev.filter((n) => n !== name) : [...prev, name]));
  };

  // No overarching card title or intro paragraph — THREE distinct labeled sub-sections below
  // already say everything: "TABLES" (the multiselect — every table that EXISTS), "SELECTED
  // TABLES" (the chip row — every table you've actually picked, click a chip to expand/collapse
  // its columns), "DIMENSIONS" (the picked-and-expanded tables' actual draggable columns). A
  // generic "Available fields" heading above all three read as one more, redundant label once
  // those three already existed (flagged live, 2026-09-27 — removed on request).
  return (
    <div className="card">
      <p className="mb-1.5 text-xs font-bold uppercase tracking-wide text-muted">Tables</p>
      {/* tablesLoading: just the table NAMES (302 of them) — fast, no column join, so this is
          brief. fieldsLoading: columns for whichever tables are newly picked — see PivotBuilder's
          own WHY comment on why this is fetched lazily per table instead of all 2301 up front. */}
      {tablesLoading ? (
        <p className="mb-2 flex items-center gap-1.5 text-xs text-muted">
          <Loader2 className="h-3.5 w-3.5 animate-spin" />
          Loading the list of tables…
        </p>
      ) : (
        <>
          <MultiSelectField
            options={tables.map((t) => t.displayName)}
            value={enabledTableNames}
            onChange={onEnabledTableNamesChange}
            placeholder="Search and pick tables…"
          />
          {fieldsLoading && (
            <p className="mb-2 flex items-center gap-1.5 text-xs text-muted">
              <Loader2 className="h-3.5 w-3.5 animate-spin" />
              Loading columns…
            </p>
          )}
          {enabledTableNames.length === 0 ? (
            <p className="mb-2 text-xs text-muted">Pick at least one table above to see its columns here.</p>
          ) : (
            <>
              <p className="mb-1.5 mt-3 text-xs font-bold uppercase tracking-wide text-muted">Selected tables</p>
              <ToggleChipRow
                items={groups.map((g) => ({ name: g.name, badge: g.fields.length }))}
                visibleNames={visibleGroupNames}
                onToggle={toggleGroup}
                onSelectAll={() => setVisibleGroupNames(groups.map((g) => g.name))}
                onClearAll={() => setVisibleGroupNames([])}
              />
            </>
          )}
        </>
      )}
      {/* Dimensions label only once something's actually expanded — no point heading an empty
          grid. items-start: without it, CSS grid stretches every card in a row to match the row's
          TALLEST sibling (e.g. a 10-field table forcing a 2-field table's card to the same
          height) — each card should only be as tall as its own field list, up to its own scroll
          cap (see the <ul>'s own comment below) rather than a max-height on the WHOLE grid, which
          used to scroll every card together — including the mostly-empty ones — just because ONE
          table's list was long (flagged live, 2026-09-27, with a single 33-field table). */}
      {visibleGroups.length > 0 && <p className="mb-1.5 mt-3 text-xs font-bold uppercase tracking-wide text-muted">Dimensions</p>}
      <div className="grid grid-cols-1 items-start gap-3 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
        {visibleGroups.map((group) => (
          <div key={group.name} className="overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm">
            {/* Solid gold, same as every table's chip above — no more one-of-10 rainbow color per
                table (see ToggleChipRow's own WHY comment). */}
            <div className="flex items-center justify-between gap-2 bg-gold px-3 py-1.5">
              <span className="truncate text-xs font-bold uppercase tracking-wide text-white">{group.name}</span>
              <span className="shrink-0 rounded-full bg-white/25 px-2 py-0.5 text-[10px] font-bold text-white">{group.fields.length}</span>
            </div>
            {/* Scrolls THIS table's own list past 30rem, instead of the old max-height on the whole
                grid above scrolling every visible card together over one long table. */}
            <ul className="max-h-[30rem] overflow-y-auto">
              {group.fields.map(({ field, label }, fieldIndex) => (
                <li
                  key={field}
                  draggable
                  onDragStart={(e) => setFieldDragPayload(e, field, "available")}
                  className={`flex cursor-grab items-center gap-1.5 px-3 py-1.5 font-mono text-[11.5px] text-ink hover:bg-slate-50 active:cursor-grabbing ${
                    fieldIndex > 0 ? "border-t border-slate-100" : ""
                  }`}
                >
                  <span className={`h-1.5 w-1.5 shrink-0 rounded-full ${isPlaced(field) ? "bg-gold" : "bg-transparent"}`} aria-hidden="true" />
                  <span className="truncate">{label}</span>
                </li>
              ))}
            </ul>
          </div>
        ))}
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
// emptyIsDropTarget: false for the Values zone (its only caller with this prop set) — Values
// never accepts a NEW field by dropping into its empty state (measures are added via the
// multi-select above, see AvailableFieldsPanel's WHY comment), so the dashed "you can drop here"
// box was misleading there — flagged live as confusing ("zero functional value") since it visually
// promised drag-and-drop that doesn't apply to an empty Values zone. Filters/Rows/Columns keep the
// real drop-target styling (dragging IS how you add a field there).
function ZoneDropArea({ zoneKey, onFieldDropped, isFieldAccepted, isEmpty, emptyText = "Drag a field here", emptyIsDropTarget = true, children }) {
  const [dragOver, setDragOver] = useState(false);

  const handleDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    const payload = readFieldDragPayload(e);
    if (payload && payload.source !== zoneKey && (!isFieldAccepted || isFieldAccepted(payload.field))) {
      onFieldDropped(payload.field, payload.source, zoneKey);
    }
  };

  return (
    <div
      className={`flex-1 rounded-lg ${dragOver ? "bg-gold/10 ring-2 ring-gold ring-inset" : ""} ${
        isEmpty && emptyIsDropTarget ? "min-h-[2rem] border-2 border-dashed border-slate-200 p-2" : isEmpty ? "" : ""
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

const FILTER_MODE_OPTIONS = [
  { value: "Members", label: "Members (pick exact values)" },
  { value: "Exclude", label: "Exclude (hide picked values)" },
  { value: "LabelContains", label: "Label contains…" },
  { value: "LabelBeginsWith", label: "Label begins with…" },
  { value: "LabelEndsWith", label: "Label ends with…" },
  { value: "TopN", label: "Top N by value" },
  { value: "BottomN", label: "Bottom N by value" },
  { value: "GreaterThan", label: "Value greater than…" },
  { value: "LessThan", label: "Value less than…" },
  { value: "Between", label: "Value between…" },
];

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
function FilterRow({ filter, dimensionOptions, selectedValueOptions, onChange, onRemove }) {
  const [memberOptions, setMemberOptions] = useState([]);
  const [loadingMembers, setLoadingMembers] = useState(false);
  const [memberSearch, setMemberSearch] = useState("");
  const mode = filter.mode ?? "Members";

  // A search typed for the PREVIOUS field shouldn't silently carry over and narrow the new
  // field's member list before the user has typed anything for it.
  useEffect(() => {
    setMemberSearch("");
  }, [filter.field]);

  useEffect(() => {
    if (!filter.field || (mode !== "Members" && mode !== "Exclude")) {
      setMemberOptions([]);
      return;
    }
    let cancelled = false;
    setLoadingMembers(true);
    // Debounced: waits for a pause in typing instead of firing one cube query per keystroke,
    // same reasoning as CardinalityGuard exists at all — the shared cube server has crashed
    // before under query load.
    const timeoutId = setTimeout(() => {
      getMembers(filter.field, memberSearch || undefined)
        .then((members) => {
          if (!cancelled) setMemberOptions(members);
        })
        .finally(() => {
          if (!cancelled) setLoadingMembers(false);
        });
    }, 400);
    return () => {
      cancelled = true;
      clearTimeout(timeoutId);
    };
  }, [filter.field, mode, memberSearch]);

  const selectedField = dimensionOptions.find((opt) => opt.value === filter.field) ?? null;

  const handleModeChange = (nextMode) => {
    onChange({ field: filter.field, mode: nextMode, includedMembers: [], labelText: "", n: 10, byMeasureField: null, value: null, valueTo: null });
  };

  return (
    <div className="mb-4 rounded-xl border border-slate-200 bg-slate-50/80 p-4">
      <div className="mb-3 flex items-start gap-3">
        {filter.field && (
          <span
            draggable
            onDragStart={(e) => setFieldDragPayload(e, filter.field, "filters")}
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
            onChange={(opt) => onChange({ field: opt?.value ?? "", mode: "Members", includedMembers: [] })}
            isClearable
            menuPortalTarget={typeof document !== "undefined" ? document.body : null}
            menuPosition="fixed"
            styles={{ menuPortal: (base) => ({ ...base, zIndex: 10000 }) }}
          />
        </div>
        <button type="button" onClick={onRemove} className="mt-2 text-muted hover:text-ink" aria-label="Remove filter">
          <X className="h-5 w-5" />
        </button>
      </div>
      {filter.field && (
        <div className="flex flex-col gap-3">
          <select className={nativeSelectClass} value={mode} onChange={(e) => handleModeChange(e.target.value)}>
            {FILTER_MODE_OPTIONS.map((opt) => (
              <option key={opt.value} value={opt.value}>
                {opt.label}
              </option>
            ))}
          </select>

          {(mode === "Members" || mode === "Exclude") && (
            <MultiSelectField
              options={memberOptions}
              value={filter.includedMembers}
              onChange={(includedMembers) => onChange({ ...filter, includedMembers })}
              placeholder={
                loadingMembers ? "Loading members…" : mode === "Exclude" ? "Type to search values to exclude…" : "Type to search members…"
              }
              isLoading={loadingMembers}
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

          {(mode === "LabelContains" || mode === "LabelBeginsWith" || mode === "LabelEndsWith") && (
            <input
              type="text"
              className="input-field"
              placeholder="Text to match…"
              value={filter.labelText ?? ""}
              onChange={(e) => onChange({ ...filter, labelText: e.target.value })}
            />
          )}

          {(mode === "TopN" || mode === "BottomN") && (
            <div className="flex flex-wrap items-center gap-3">
              <input
                type="number"
                min={1}
                className="input-field w-24"
                value={filter.n ?? 10}
                onChange={(e) => onChange({ ...filter, n: Number(e.target.value) })}
              />
              <select
                className={nativeSelectClass}
                value={filter.byMeasureField ?? ""}
                onChange={(e) => onChange({ ...filter, byMeasureField: e.target.value || null })}
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
                value={filter.value ?? ""}
                onChange={(e) => onChange({ ...filter, value: e.target.value === "" ? null : Number(e.target.value) })}
              />
              {mode === "Between" && (
                <input
                  type="number"
                  className="input-field w-28"
                  placeholder="To…"
                  value={filter.valueTo ?? ""}
                  onChange={(e) => onChange({ ...filter, valueTo: e.target.value === "" ? null : Number(e.target.value) })}
                />
              )}
              <select
                className={nativeSelectClass}
                value={filter.byMeasureField ?? ""}
                onChange={(e) => onChange({ ...filter, byMeasureField: e.target.value || null })}
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
    </div>
  );
}

// The compact Filters drop-zone box shows just this pill (field name only) — the full config
// (mode, members/text/N, ...) lives in a Popover behind it, reusing FilterRow unchanged. The pill
// itself is draggable so moving a filter to another zone doesn't require opening the popover
// first, same pattern as HierarchicalCubePathSelect's own pills.
function FilterPill({ filter, dimensionOptions, selectedValueOptions, onChange, onRemove }) {
  const label = filter.field ? formatDimensionName(filter.field) : "New filter";
  return (
    <Popover
      trigger={(toggle) => (
        <span
          draggable={Boolean(filter.field)}
          onDragStart={(e) => filter.field && setFieldDragPayload(e, filter.field, "filters")}
          className={`inline-flex items-center gap-1 rounded-md bg-gold px-2 py-1 text-xs font-medium text-white ${
            filter.field ? "cursor-grab active:cursor-grabbing" : ""
          }`}
        >
          <button type="button" onClick={toggle} className="cursor-pointer">
            {label}
          </button>
          <button type="button" onClick={onRemove} className="rounded p-0.5 hover:bg-white/20" aria-label="Remove filter">
            <X className="h-3 w-3" />
          </button>
        </span>
      )}
    >
      {() => (
        <FilterRow
          filter={filter}
          dimensionOptions={dimensionOptions}
          selectedValueOptions={selectedValueOptions}
          onChange={onChange}
          onRemove={onRemove}
        />
      )}
    </Popover>
  );
}

// Excel's "Group Field" (date) and "Group Selection" (numeric bins), one control per Rows OR
// Columns field (2026-09-24: extended from Rows-only) — backend enforces one grouping per field
// (see MdxPivotQueryBuilder.ValidateGroupings). Mode is derived from whichever of
// dateGroupings/numericGroupings currently mentions the field, rather than stored as its own
// separate piece of state, so the two arrays stay the single source of truth the backend actually
// reads. Rendered once per axis (see the two call sites below) — dateGroupings/numericGroupings
// are shared arrays across both, since a grouping object only names its field, not an axis.
function AxisGroupingEditor({ label, fields, dateGroupings, numericGroupings, onChange }) {
  if (fields.length === 0) {
    return null;
  }

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
      {fields.map((field) => {
        const mode = modeFor(field);
        return (
          <div key={field} className="flex flex-wrap items-center gap-2">
            <span className="min-w-[10rem] text-sm text-ink">{formatDimensionName(field)}</span>
            <select className={nativeSelectClass} value={mode} onChange={(e) => setMode(field, e.target.value)}>
              <option value="None">No grouping</option>
              <option value="Month">Group by Month</option>
              <option value="Quarter">Group by Quarter</option>
              <option value="Year">Group by Year</option>
              <option value="Numeric">Numeric bins</option>
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

  useEffect(() => {
    if (!field) {
      setMemberOptions([]);
      return;
    }
    let cancelled = false;
    setLoadingMembers(true);
    getMembers(field)
      .then((members) => {
        if (!cancelled) setMemberOptions(members);
      })
      .finally(() => {
        if (!cancelled) setLoadingMembers(false);
      });
    return () => {
      cancelled = true;
    };
  }, [field]);

  if (!field) {
    return <p className="text-sm text-muted">Add exactly one field to the target axis first to set a custom order.</p>;
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
function SortEditor({ sort, rowLevelSorts, columnLevelSorts, values, rows, columns, onChange }) {
  const mode = sort === null ? "none" : sort.customOrder ? "custom" : sort.byMeasureField ? sort.byMeasureField : "label";
  const axis = sort?.axis ?? "Rows";
  const targetField = (axis === "Columns" ? columns : rows)[0] ?? null;

  const handleModeChange = (nextMode) => {
    // Setting the single axis-wide sort clears any level-by-level rules on the SAME axis — the
    // backend rejects having both at once (ValidateLevelSorts). The other axis's level sorts are
    // untouched (they're independent of this axis's own Sort field).
    const clearLevelSorts = axis === "Columns" ? { columnLevelSorts: [] } : { rowLevelSorts: [] };
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
function CalculatedFieldsEditor({ calculatedFields, measureOptions, onChange }) {
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

  useEffect(() => {
    if (!item.field) {
      setMemberOptions([]);
      return;
    }
    let cancelled = false;
    setLoadingMembers(true);
    getMembers(item.field)
      .then((members) => {
        if (!cancelled) setMemberOptions(members);
      })
      .finally(() => {
        if (!cancelled) setLoadingMembers(false);
      });
    return () => {
      cancelled = true;
    };
  }, [item.field]);

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
      {item.field && (
        <div className="flex flex-col gap-2">
          <div>
            <span className="mb-1 block text-xs font-medium text-muted">Add (+):</span>
            <MultiSelectField
              options={memberOptions}
              value={item.positiveMembers}
              onChange={(positiveMembers) => onChange({ positiveMembers })}
              placeholder={loadingMembers ? "Loading members…" : "Select member(s) to add…"}
              isLoading={loadingMembers}
            />
          </div>
          <div>
            <span className="mb-1 block text-xs font-medium text-muted">Subtract (−):</span>
            <MultiSelectField
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

function CalculatedItemsEditor({ calculatedItems, rows, columns, onChange }) {
  const axisFieldOptions = [...rows, ...columns];

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

// A CalculatedItem's Field must still be a Rows or Columns field — the backend rejects one that
// isn't (see MdxPivotQueryBuilder.ValidateCalculatedItems) — so it needs to be dropped the moment
// its field leaves both axes, same idea as dateGroupings/numericGroupings being dropped when their
// field leaves Rows.
const clearOrphanedCalculatedItems = (calculatedItems, rows, columns) =>
  calculatedItems.filter((item) => rows.includes(item.field) || columns.includes(item.field));

// The compact Values drop-zone box shows just this pill (measure name only) — which measure it is
// gets set once, from the multi-select above the pill list (see FieldPicker's own render); the
// pill itself only carries per-measure CONFIG (Show Values As, Format, Conditional format) behind
// a Popover, plus drag-to-reorder within the Values zone (onDropReorder), same pattern as
// FilterPill.
function ValuePill({ v, onUpdate, onUpdateFormat, onRemove, onDropReorder, rows }) {
  const format = v.format ?? DEFAULT_FORMAT;
  const isPercentOf = v.showValuesAs === "PercentOf";
  const [baseItemOptions, setBaseItemOptions] = useState([]);
  const [loadingBaseItems, setLoadingBaseItems] = useState(false);

  // Base item options come from the SAME member-search endpoint FilterRow uses — only fetched once
  // a Base field is actually chosen, and re-fetched whenever it changes.
  useEffect(() => {
    if (!isPercentOf || !v.baseField) {
      setBaseItemOptions([]);
      return;
    }
    let cancelled = false;
    setLoadingBaseItems(true);
    getMembers(v.baseField)
      .then((members) => {
        if (!cancelled) setBaseItemOptions(members);
      })
      .finally(() => {
        if (!cancelled) setLoadingBaseItems(false);
      });
    return () => {
      cancelled = true;
    };
  }, [isPercentOf, v.baseField]);

  return (
    <div onDragOver={(e) => e.preventDefault()} onDrop={onDropReorder}>
      {/* align="right": the Values card sits flush against the right edge of the "Fields"
          strip (it's a w-fit card pushed there by Filters/Columns/Rows' flex-1 — see FieldPicker's
          own WHY comment on that grid), so Popover's default left-aligned/expand-right panel had
          nowhere to expand into and was clipped by the browser's edge (caught live, 2026-09-27).
          Opening it expanding leftward, into the room Filters/Columns/Rows have, keeps it on
          screen. */}
      <Popover
        align="right"
        trigger={(toggle) => (
          <span
            draggable
            onDragStart={(e) => setFieldDragPayload(e, v.field, "values")}
            className="inline-flex cursor-grab items-center gap-1 rounded-md bg-gold px-2 py-1 text-xs font-medium text-white active:cursor-grabbing"
          >
            <button type="button" onClick={toggle} className="cursor-pointer">
              {v.field}
            </button>
            <button type="button" onClick={onRemove} className="rounded p-0.5 hover:bg-white/20" aria-label="Remove value">
              <X className="h-3 w-3" />
            </button>
          </span>
        )}
      >
        {() => (
          <div className="space-y-3">
            <div>
              <span className="mb-1 block text-xs font-medium text-muted">Show values as</span>
              <select className={nativeSelectClass} value={v.showValuesAs} onChange={(e) => onUpdate({ showValuesAs: e.target.value })}>
                {SHOW_VALUES_AS_OPTIONS.map((opt) => (
                  <option key={opt.value} value={opt.value}>
                    {opt.label}
                  </option>
                ))}
              </select>
            </div>
            {isPercentOf && (
              <div className="space-y-2 rounded-md bg-slate-50 p-2">
                <div>
                  <span className="mb-1 block text-xs font-medium text-muted">Base field (Rows)</span>
                  <select
                    className={nativeSelectClass}
                    value={v.baseField ?? ""}
                    onChange={(e) => onUpdate({ baseField: e.target.value, baseItem: "" })}
                  >
                    <option value="">Choose a row field…</option>
                    {rows.map((field) => (
                      <option key={field} value={field}>
                        {formatDimensionName(field)}
                      </option>
                    ))}
                  </select>
                </div>
                {v.baseField && (
                  <div>
                    <span className="mb-1 block text-xs font-medium text-muted">Base item</span>
                    <select
                      className={nativeSelectClass}
                      value={v.baseItem ?? ""}
                      onChange={(e) => onUpdate({ baseItem: e.target.value })}
                      disabled={loadingBaseItems}
                    >
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
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-xs font-medium text-muted">Format:</span>
              <select className={nativeSelectClass} value={format.type} onChange={(e) => onUpdateFormat({ type: e.target.value })}>
                <option value="General">General</option>
                <option value="Number">Number</option>
                <option value="Currency">Currency</option>
                <option value="Percentage">Percentage</option>
              </select>
              {format.type !== "General" && (
                <input
                  type="number"
                  min={0}
                  max={10}
                  value={format.decimalPlaces}
                  onChange={(e) => onUpdateFormat({ decimalPlaces: Number(e.target.value) })}
                  className="input-field w-20 py-2"
                  title="Decimal places"
                />
              )}
              {format.type === "Currency" && (
                <input
                  type="text"
                  value={format.currencySymbol}
                  onChange={(e) => onUpdateFormat({ currencySymbol: e.target.value })}
                  placeholder="EGP"
                  className="input-field w-24 py-2"
                  title="Currency symbol"
                />
              )}
            </div>
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-xs font-medium text-muted">Conditional format:</span>
              <select
                className={nativeSelectClass}
                value={v.conditionalFormat?.type ?? "None"}
                onChange={(e) => onUpdate({ conditionalFormat: { type: e.target.value } })}
              >
                <option value="None">None</option>
                <option value="ColorScale">Color scale</option>
                <option value="DataBar">Data bar</option>
                <option value="IconSet">Icon set (traffic lights)</option>
              </select>
            </div>
          </div>
        )}
      </Popover>
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
}) {
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

  const updateValue = (index, patch) => {
    const copy = [...value.values];
    copy[index] = { ...copy[index], ...patch };
    onChange({ ...value, values: copy });
  };

  const updateValueFormat = (index, patch) => {
    const current = value.values[index].format ?? DEFAULT_FORMAT;
    updateValue(index, { format: { ...current, ...patch } });
  };

  const removeValue = (index) => {
    onChange({ ...value, values: value.values.filter((_, i) => i !== index) });
  };

  // The single add/remove entry point for Values, driven by the multi-select above the pill list
  // (see the "Values" box below) — a plain diff against the current pills: fields no longer
  // selected are dropped, newly-selected fields become new pills, and anything unchanged KEEPS
  // its existing config (showValuesAs/format/conditionalFormat) rather than resetting it, since
  // re-selecting the same measure isn't meant to wipe out per-measure settings already made.
  const setSelectedValueFields = (selectedFields) => {
    const kept = value.values.filter((v) => selectedFields.includes(v.field));
    const added = selectedFields
      .filter((field) => !value.values.some((v) => v.field === field))
      .map((field) => ({ field, showValuesAs: "Normal" }));
    onChange({ ...value, values: [...kept, ...added] });
  };

  const updateFilter = (index, patch) => {
    const copy = [...value.filters];
    copy[index] = patch;
    onChange({ ...value, filters: copy });
  };

  const removeFilter = (index) => {
    onChange({ ...value, filters: value.filters.filter((_, i) => i !== index) });
  };

  const addFilter = () => {
    onChange({ ...value, filters: [...value.filters, { field: "", includedMembers: [] }] });
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

    if (toZone === "rows" || toZone === "columns" || toZone === "values") {
      next.rows = value.rows.filter((f) => f !== field);
      next.columns = value.columns.filter((f) => f !== field);
      next.values = value.values.filter((v) => v.field !== field);
      // A field leaving Rows takes its grouping with it — the backend rejects a grouping on a
      // field that's no longer in Rows (ValidateGroupings).
      next.dateGroupings = value.dateGroupings.filter((g) => g.field !== field);
      next.numericGroupings = value.numericGroupings.filter((g) => g.field !== field);
    }

    if (toZone === "rows" && !next.rows.includes(field)) {
      next.rows = [...next.rows, field];
    } else if (toZone === "columns" && !next.columns.includes(field)) {
      next.columns = [...next.columns, field];
    } else if (toZone === "values" && !next.values.some((v) => v.field === field)) {
      next.values = [...next.values, { field, showValuesAs: "Normal" }];
    } else if (toZone === "filters" && !next.filters.some((f) => f.field === field)) {
      next.filters = [...next.filters, { field, mode: "Members", includedMembers: [] }];
    }

    next.calculatedItems = clearOrphanedCalculatedItems(value.calculatedItems, next.rows, next.columns);

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

  const isFieldPlaced = (field) =>
    value.rows.includes(field) || value.columns.includes(field) || value.values.some((v) => v.field === field) || value.filters.some((f) => f.field === field);

  return (
    <div className="grid grid-cols-1 gap-6">
      <AvailableFieldsPanel
        dimensions={dimensions}
        isPlaced={isFieldPlaced}
        tables={tables}
        tablesLoading={tablesLoading}
        enabledTableNames={enabledTableNames}
        onEnabledTableNamesChange={onEnabledTableNamesChange}
        fieldsLoading={fieldsLoading}
      />

      {/* Compact fields strip — Filters/Columns/Rows/Values always visible as small boxes, each
          with its own pill list. A pill shows only its field/measure name; the full add/edit form
          (dimension picker, filter mode, show-values-as, ...) lives in a Popover behind the box's
          "+" or, for Filters/Values, behind the pill itself — one click away instead of
          permanently occupying space, per-item since a filter/value pill needs its OWN config, not
          just a name. Labeled "Fields", not "Drop zones": Values doesn't actually accept a field
          dropped onto it — measures have no drag source anywhere (added only via its "Choose
          measures…" select, see AvailableFieldsPanel's own WHY comment) — only Filters/Columns/Rows
          are real drop targets, so calling all four "drop zones" was misleading (flagged live,
          2026-09-27). */}
      <div>
        {/* Reset sits at the right of this row, roughly above the Values card below (the
            rightmost of the four) — moved out of the page header per explicit request,
            2026-09-27. */}
        <div className="mb-2 flex items-center justify-between">
          <p className="text-xs font-bold uppercase tracking-wide text-muted">Fields</p>
          <button type="button" onClick={onReset} className="btn-secondary px-3 py-1.5 text-xs">
            Reset
          </button>
        </div>
        {/* flex, not grid: Values is a fixed w-fit card (it only needs enough width for its
            "Values on: ○ Columns ○ Rows" line — see its own comment below), and a 4-column grid
            left the rest of that track empty. flex-1 on the other three lets them actually claim
            that freed width instead of leaving blank space (flagged live from a screenshot,
            2026-09-27). */}
        <div className="flex flex-wrap items-stretch gap-2">
          {/* Filters/Columns/Rows share ONE card now (requested live, 2026-09-27, right after
              renaming this strip's own label from "Drop zones" to "Fields" for the same reason):
              they're the three fields that actually accept a dropped field, so grouping them into
              one visual card and leaving Values — which never accepts one — as its own separate
              card makes that distinction visible, not just documented in a comment. divide-x draws
              the seam between them instead of three separate borders/shadows. No h-full on this
              outer card: "height: 100%" computes to a non-auto value, which per the flexbox spec
              DISABLES align-self:stretch (stretch only overrides a cross-size that computes to
              auto) — caught live the same day on the old separate-cards layout, see displayValues'
              sibling fix in PivotBuilder.jsx for the analogous bug. */}
          <div className="flex min-w-[220px] flex-1 items-stretch divide-x divide-slate-200 self-stretch rounded-xl border border-slate-200 bg-white shadow-sm">
            <div className="flex min-w-[180px] flex-1 flex-col p-2">
              <div className="mb-1.5 flex items-center justify-between gap-2">
                <span className="flex items-center gap-1.5 text-xs font-bold uppercase tracking-wide text-muted">
                  <Filter className="h-3.5 w-3.5" />
                  Filters {value.filters.length > 0 && <span className="text-gold">({value.filters.length})</span>}
                </span>
                <button
                  type="button"
                  onClick={addFilter}
                  className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-gold/15 text-gold hover:bg-gold/25"
                  aria-label="Add filter"
                  title="Add filter"
                >
                  <Plus className="h-3.5 w-3.5" />
                </button>
              </div>
              <ZoneDropArea zoneKey="filters" onFieldDropped={moveField} isEmpty={value.filters.length === 0}>
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
                    />
                  ))}
                </div>
              </ZoneDropArea>
            </div>

            <div className="min-w-[180px] flex-1">
              <HierarchicalCubePathSelect
                bare
                label="Columns"
                icon={Columns3}
                options={dimensions.map((d) => d.field)}
                value={value.columns}
                onChange={(columns) => {
                  // Dropping a column field also drops any grouping/calculated item that referenced
                  // it — the backend rejects both on a field that's no longer in Rows/Columns
                  // (ValidateGroupings, ValidateCalculatedItems). A field still present in Rows keeps
                  // its grouping (it didn't actually leave either axis).
                  const stillGrouped = (field) => columns.includes(field) || value.rows.includes(field);
                  const dateGroupings = value.dateGroupings.filter((g) => stillGrouped(g.field));
                  const numericGroupings = value.numericGroupings.filter((g) => stillGrouped(g.field));
                  const calculatedItems = clearOrphanedCalculatedItems(value.calculatedItems, value.rows, columns);
                  // columnLevelSorts' `level` values are positional indexes into Columns — any add,
                  // remove, or reorder here can shift which field a stored index actually points to,
                  // so they're cleared rather than risking a rule silently applying to the wrong
                  // field (same defensive-clear approach as groupings/calculatedItems above).
                  onChange({ ...value, columns, dateGroupings, numericGroupings, calculatedItems, columnLevelSorts: [] });
                }}
                dragSourceKey="columns"
                onFieldDropped={moveField}
              />
            </div>

            <div className="min-w-[180px] flex-1">
              <HierarchicalCubePathSelect
                bare
                align="right"
                label="Rows"
                icon={Rows3}
                options={dimensions.map((d) => d.field)}
                value={value.rows}
                onChange={(rows) => {
                  // Mirror of the Columns handler above — a field still present in Columns keeps its
                  // grouping.
                  const stillGrouped = (field) => rows.includes(field) || value.columns.includes(field);
                  const dateGroupings = value.dateGroupings.filter((g) => stillGrouped(g.field));
                  const numericGroupings = value.numericGroupings.filter((g) => stillGrouped(g.field));
                  const calculatedItems = clearOrphanedCalculatedItems(value.calculatedItems, rows, value.columns);
                  // Mirror of the columnLevelSorts clear above.
                  onChange({ ...value, rows, dateGroupings, numericGroupings, calculatedItems, rowLevelSorts: [] });
                }}
                dragSourceKey="rows"
                onFieldDropped={moveField}
              />
            </div>
          </div>

          {/* No separate aggregation dropdown: confirmed against the real cube (2026-09-21) that
              measures are pre-built with their aggregation baked in — see PivotValueField in the
              backend for the full story. Measures are ADDED here via the multi-select (a measure
              has exactly one valid destination, so there's nothing for drag to decide — see
              AvailableFieldsPanel's WHY comment); a pill's own Popover then covers Show Values
              As/Format/Conditional format, and pills can still be drag-reordered among themselves
              since that order affects the generated MDX (see handleValueRowDrop). */}
          {/* justify-self-start + w-fit: the grid track this sits in is as wide as the other 3
              cards, but this card's own content (a short select + two radio labels) doesn't need
              that width — stretching it to fill the track just left a lot of empty space on the
              right (flagged live from a screenshot, 2026-09-27). w-fit rather than a guessed fixed
              width so it hugs the "Values on: ○ Columns ○ Rows" line exactly, whatever that line's
              actual rendered width is — a hardcoded w-56 was narrower than that line and wrapped
              "Values on:" onto two lines (also caught live). */}
          {/* No h-full here either — see the Filters card's WHY comment above; same fix. */}
          <div className="flex w-fit flex-none flex-col self-stretch rounded-xl border border-slate-200 bg-white p-2 shadow-sm">
            <div className="mb-1.5 flex items-center justify-between gap-2">
              <span className="flex items-center gap-1.5 text-xs font-bold uppercase tracking-wide text-muted">
                <Sigma className="h-3.5 w-3.5" />
                Values {value.values.length > 0 && <span className="text-gold">({value.values.length})</span>}
              </span>
            </div>
            <Select
              isMulti
              className="react-select-container mb-1.5"
              classNamePrefix="react-select"
              placeholder="Choose measures…"
              options={valueFieldOptions}
              value={valueFieldOptions.filter((opt) => value.values.some((v) => v.field === opt.value))}
              onChange={(selected) => setSelectedValueFields((selected ?? []).map((opt) => opt.value))}
              menuPortalTarget={typeof document !== "undefined" ? document.body : null}
              menuPosition="fixed"
              styles={{
                menuPortal: (base) => ({ ...base, zIndex: 10000 }),
                // Compact card, not react-select's default ~38px control — shrunk alongside the
                // rest of the Drop Zones strip (caught live, 2026-09-27: the strip was still too
                // tall after the "Values on" line-break fix, because this control's default
                // height was the next biggest contributor).
                control: (base) => ({ ...base, minHeight: "30px", minWidth: "180px" }),
                valueContainer: (base) => ({ ...base, padding: "0 6px" }),
                indicatorsContainer: (base) => ({ ...base, height: "30px" }),
              }}
            />
            {/* role="radiogroup" instead of a real <fieldset>/<legend>: a <legend> is block-level
                and forces its own line above the radios, which — via the "Fields" strip's
                items-stretch — was making every OTHER card stretch taller to match (caught live,
                2026-09-27). A single-line label + radios keeps the same height this box had with
                the old <select>. */}
            <div
              className="mb-1.5 flex items-center gap-3 whitespace-nowrap"
              role="radiogroup"
              aria-label="Values on"
              title="Excel's own 'drag Σ Values between Columns and Rows' — see DECISIONS.md for this phase's scope cuts (needs a Column field, Normal-only Show Values As, no chart)"
            >
              <span className="text-xs font-medium text-muted">Values on:</span>
              {["Columns", "Rows"].map((placement) => (
                <label key={placement} className="flex items-center gap-1.5 text-sm text-ink">
                  <input
                    type="radio"
                    name="valuesPlacement"
                    value={placement}
                    checked={(value.valuesPlacement ?? "Columns") === placement}
                    onChange={(e) => onChange({ ...value, valuesPlacement: e.target.value })}
                    className="h-4 w-4 accent-gold"
                  />
                  {placement}
                </label>
              ))}
            </div>
            <ZoneDropArea
              zoneKey="values"
              onFieldDropped={moveField}
              isFieldAccepted={(field) => valueFieldOptions.some((opt) => opt.value === field)}
              isEmpty={value.values.length === 0}
              emptyText=""
              emptyIsDropTarget={false}
            >
              <div className="flex flex-wrap gap-1.5">
                {value.values.map((v, index) => (
                  <ValuePill
                    key={index}
                    v={v}
                    rows={value.rows}
                    onUpdate={(patch) => updateValue(index, patch)}
                    onUpdateFormat={(patch) => updateValueFormat(index, patch)}
                    onRemove={() => removeValue(index)}
                    onDropReorder={(e) => handleValueRowDrop(e, index)}
                  />
                ))}
              </div>
            </ZoneDropArea>
          </div>
        </div>

        <AxisGroupingEditor
          label="Group Rows by"
          fields={value.rows}
          dateGroupings={value.dateGroupings}
          numericGroupings={value.numericGroupings}
          onChange={(dateGroupings, numericGroupings) => onChange({ ...value, dateGroupings, numericGroupings })}
        />
        <AxisGroupingEditor
          label="Group Columns by"
          fields={value.columns}
          dateGroupings={value.dateGroupings}
          numericGroupings={value.numericGroupings}
          onChange={(dateGroupings, numericGroupings) => onChange({ ...value, dateGroupings, numericGroupings })}
        />
      </div>

      {/* Secondary, optional settings — small pill controls instead of full-width cards, each
          opening a Popover with its existing (unchanged) editor. items-start, not items-center:
          each Popover below is `inline` now (its panel renders in normal flow under its own
          button, pushing "Run query" down instead of floating over it) — with items-center, one
          button's column growing taller than its siblings' would re-center ALL of them against the
          new row height, making the untouched buttons visibly drift down/float instead of staying
          put at the top (caught live, 2026-09-27). items-start keeps every trigger button flush at
          the row's top regardless of which one panel, if any, is open. */}
      <div className="flex flex-wrap items-start gap-2">
        <Popover
          inline
          trigger={(toggle) => (
            <button type="button" onClick={toggle} className="btn-secondary inline-flex items-center gap-1.5 px-3 py-1.5 text-xs">
              Sort
              {(value.sort || value.rowLevelSorts.length > 0 || value.columnLevelSorts.length > 0) && (
                <span className="rounded-full bg-gold/20 px-1.5 text-gold">
                  {value.sort ? 1 : value.rowLevelSorts.length + value.columnLevelSorts.length}
                </span>
              )}
              <ChevronDown className="h-3.5 w-3.5" />
            </button>
          )}
        >
          {() => (
            <SortEditor
              sort={value.sort}
              rowLevelSorts={value.rowLevelSorts}
              columnLevelSorts={value.columnLevelSorts}
              values={value.values}
              rows={value.rows}
              columns={value.columns}
              onChange={(patch) => onChange({ ...value, ...patch })}
            />
          )}
        </Popover>

        <Popover
          inline
          trigger={(toggle) => (
            <button type="button" onClick={toggle} className="btn-secondary inline-flex items-center gap-1.5 px-3 py-1.5 text-xs">
              Totals
              {(value.showGrandTotals || value.showSubtotals) && (
                <span className="rounded-full bg-gold/20 px-1.5 text-gold">
                  {(value.showGrandTotals ? 1 : 0) + (value.showSubtotals ? 1 : 0)}
                </span>
              )}
              <ChevronDown className="h-3.5 w-3.5" />
            </button>
          )}
        >
          {() => (
            <div>
              <label className="mb-2 flex items-center gap-2 text-sm font-medium text-ink">
                <input
                  type="checkbox"
                  className="h-4 w-4 accent-gold"
                  checked={value.showGrandTotals}
                  onChange={(e) => onChange({ ...value, showGrandTotals: e.target.checked })}
                />
                Show Grand Totals
              </label>
              <label className="flex items-center gap-2 text-sm font-medium text-ink">
                <input
                  type="checkbox"
                  className="h-4 w-4 accent-gold"
                  checked={value.showSubtotals}
                  disabled={value.rows.length < 2}
                  onChange={(e) => onChange({ ...value, showSubtotals: e.target.checked })}
                />
                Show Subtotals (needs 2+ Row fields)
              </label>
            </div>
          )}
        </Popover>

        <Popover
          inline
          trigger={(toggle) => (
            <button type="button" onClick={toggle} className="btn-secondary inline-flex items-center gap-1.5 px-3 py-1.5 text-xs">
              No-Data Items
              {value.showItemsWithNoData.length > 0 && (
                <span className="rounded-full bg-gold/20 px-1.5 text-gold">{value.showItemsWithNoData.length}</span>
              )}
              <ChevronDown className="h-3.5 w-3.5" />
            </button>
          )}
        >
          {() => (
            <div className="max-w-xs">
              <p className="mb-2 text-xs text-muted">
                Excel's "Show items with no data" — keeps a Row/Column member visible even with zero facts, instead of hiding it. Uses more
                cube resources; the query is rejected if the estimated result is too large.
              </p>
              <label className="mb-2 flex items-center gap-2 text-sm font-medium text-ink">
                <input
                  type="checkbox"
                  className="h-4 w-4 accent-gold"
                  checked={value.showItemsWithNoData.includes("Rows")}
                  onChange={(e) =>
                    onChange({
                      ...value,
                      showItemsWithNoData: e.target.checked
                        ? [...value.showItemsWithNoData, "Rows"]
                        : value.showItemsWithNoData.filter((a) => a !== "Rows"),
                    })
                  }
                />
                Show Rows with no data
              </label>
              <label className="flex items-center gap-2 text-sm font-medium text-ink">
                <input
                  type="checkbox"
                  className="h-4 w-4 accent-gold"
                  checked={value.showItemsWithNoData.includes("Columns")}
                  onChange={(e) =>
                    onChange({
                      ...value,
                      showItemsWithNoData: e.target.checked
                        ? [...value.showItemsWithNoData, "Columns"]
                        : value.showItemsWithNoData.filter((a) => a !== "Columns"),
                    })
                  }
                />
                Show Columns with no data
              </label>
            </div>
          )}
        </Popover>

        <Popover
          inline
          trigger={(toggle) => (
            <button type="button" onClick={toggle} className="btn-secondary inline-flex items-center gap-1.5 px-3 py-1.5 text-xs">
              <Plus className="h-3.5 w-3.5" /> Calculated Field
              {value.calculatedFields.length > 0 && (
                <span className="rounded-full bg-gold/20 px-1.5 text-gold">{value.calculatedFields.length}</span>
              )}
            </button>
          )}
        >
          {() => (
            <CalculatedFieldsEditor
              calculatedFields={value.calculatedFields}
              measureOptions={measureOptions}
              onChange={(calculatedFields) => {
                // A Value referencing a calculated field becomes invalid the moment that field is
                // removed, or edited back to incomplete, AFTER already being picked in Values —
                // same dangling-reference bug as calculatedFieldOptions' own fix above, just from
                // the other direction (edited/removed after selection instead of selected before
                // completion).
                const validFieldNames = new Set(
                  measureOptions
                    .map((m) => m.value)
                    .concat(calculatedFields.filter((f) => f.name && f.leftField && f.rightField).map((f) => f.name)),
                );
                const values = value.values.filter((v) => validFieldNames.has(v.field));
                onChange({ ...value, calculatedFields, values });
              }}
            />
          )}
        </Popover>

        <Popover
          inline
          trigger={(toggle) => (
            <button type="button" onClick={toggle} className="btn-secondary inline-flex items-center gap-1.5 px-3 py-1.5 text-xs">
              <Plus className="h-3.5 w-3.5" /> Calculated Item
              {value.calculatedItems.length > 0 && (
                <span className="rounded-full bg-gold/20 px-1.5 text-gold">{value.calculatedItems.length}</span>
              )}
            </button>
          )}
        >
          {() => (
            <CalculatedItemsEditor
              calculatedItems={value.calculatedItems}
              rows={value.rows}
              columns={value.columns}
              onChange={(calculatedItems) => onChange({ ...value, calculatedItems })}
            />
          )}
        </Popover>
      </div>
    </div>
  );
}
