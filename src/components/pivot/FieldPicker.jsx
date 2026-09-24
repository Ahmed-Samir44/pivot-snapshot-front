import { useEffect, useState } from "react";
import Select from "react-select";
import { X, GripVertical, Plus, ChevronDown } from "lucide-react";
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
  { value: "Rank", label: "Rank" },
  { value: "PercentOfParentRow", label: "% of Parent Row" },
  { value: "DifferenceFrom", label: "Difference From (previous)" },
  { value: "PercentOfRowTotal", label: "% of Row Total" },
  { value: "PercentOfColumnTotal", label: "% of Column Total" },
  { value: "PercentOfParentColumn", label: "% of Parent Column" },
  { value: "Index", label: "Index" },
];

const DEFAULT_FORMAT = { type: "General", decimalPlaces: 2, currencySymbol: "EGP" };

// Same categorical palette PivotChart.jsx/PivotChartRenderer.cs already use for chart series — a
// distinct color per dimension group's card header (like the reference ERD tool's per-domain
// node colors) reusing an existing app-wide palette instead of inventing a clashing new one.
const GROUP_HEADER_COLORS = [
  "#AE8C67", "#6a7380", "#63BE7B", "#F8696B", "#FFEB84", "#638EC6", "#9a7b57", "#1f2430", "#c9a876", "#8a94a3",
];

// The palette includes a light yellow (#FFEB84) that white text would nearly vanish on — picked
// per color via relative luminance instead of assuming every entry is dark enough for white text.
function readableTextColorFor(hex) {
  const r = parseInt(hex.slice(1, 3), 16) / 255;
  const g = parseInt(hex.slice(3, 5), 16) / 255;
  const b = parseInt(hex.slice(5, 7), 16) / 255;
  const luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return luminance > 0.6 ? "#1f2430" : "#ffffff";
}

const nativeSelectClass =
  "rounded-xl border border-slate-200 bg-white/90 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent";

function SectionCard({ title, badge, children }) {
  return (
    <div className="card">
      <h3 className="mb-4 flex items-center gap-2 text-sm font-bold uppercase tracking-wide text-muted">
        {title}
        {badge > 0 && <span className="rounded-full bg-gold/20 px-2 py-0.5 text-xs font-bold text-gold">{badge}</span>}
      </h3>
      {children}
    </div>
  );
}

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
      {items.map(({ name, badge }, index) => {
        const headerColor = GROUP_HEADER_COLORS[index % GROUP_HEADER_COLORS.length];
        const textColor = readableTextColorFor(headerColor);
        const isVisible = visibleNames.includes(name);
        const badgeStyle = { background: textColor === "#ffffff" ? "rgba(255,255,255,0.3)" : "rgba(31,36,48,0.15)", color: textColor };
        return (
          <button
            key={name}
            type="button"
            onClick={() => onToggle(name)}
            className={`flex items-center gap-2.5 rounded-lg px-4 py-2.5 text-sm font-bold uppercase tracking-wide shadow-sm transition ${
              isVisible ? "ring-2 ring-ink/25 ring-offset-2" : ""
            }`}
            style={{ background: headerColor, color: textColor }}
            aria-pressed={isVisible}
            title={isVisible ? `Hide ${name}` : `Show ${name}`}
          >
            {name}
            {badge > 0 && (
              <span className="rounded-full px-1.5 py-0.5 text-[10px] font-bold leading-none" style={badgeStyle}>
                {badge}
              </span>
            )}
            <span className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full text-sm font-bold leading-none" style={badgeStyle}>
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
function AvailableFieldsPanel({ dimensions, isPlaced }) {
  const groups = [];
  const groupIndexByName = new Map();
  const addToGroup = (groupName, field, label) => {
    if (!groupIndexByName.has(groupName)) {
      groupIndexByName.set(groupName, groups.length);
      groups.push({ name: groupName, fields: [] });
    }
    groups[groupIndexByName.get(groupName)].fields.push({ field, label });
  };

  dimensions.forEach((d) => {
    const dotIndex = d.displayName.indexOf(".");
    const groupName = dotIndex === -1 ? d.displayName : d.displayName.slice(0, dotIndex);
    const label = dotIndex === -1 ? d.displayName : d.displayName.slice(dotIndex + 1);
    addToGroup(groupName, d.field, label);
  });

  // No table shown until the user explicitly picks one (+ chip or "Select all") — a real cube
  // has 8+ tables and 20+ columns, so opening with every column list expanded is exactly the
  // wall-of-content this whole redesign exists to avoid.
  const [visibleGroupNames, setVisibleGroupNames] = useState([]);
  const allGroupNames = groups.map((g) => g.name);
  const visibleGroups = groups.filter((g) => visibleGroupNames.includes(g.name));

  const toggleGroup = (name) => {
    setVisibleGroupNames((prev) => (prev.includes(name) ? prev.filter((n) => n !== name) : [...prev, name]));
  };

  return (
    <SectionCard title="Available fields">
      <p className="mb-3 text-xs text-muted">
        Pick which tables to show (+ to reveal its columns, − to hide them again), then drag a column into Filters, Rows, or Columns below. Measures are picked
        separately in the Values box.
      </p>
      <ToggleChipRow
        items={groups.map((g) => ({ name: g.name, badge: g.fields.length }))}
        visibleNames={visibleGroupNames}
        onToggle={toggleGroup}
        onSelectAll={() => setVisibleGroupNames(allGroupNames)}
        onClearAll={() => setVisibleGroupNames([])}
      />
      {/* items-start: without it, CSS grid stretches every card in a row to match the row's
          TALLEST sibling (e.g. a 10-field table forcing a 2-field table's card to the same
          height) — each card should only be as tall as its own field list. */}
      <div className="grid max-h-[30rem] grid-cols-1 items-start gap-3 overflow-y-auto sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
        {visibleGroups.map((group) => {
          const groupIndex = groups.indexOf(group);
          const headerColor = GROUP_HEADER_COLORS[groupIndex % GROUP_HEADER_COLORS.length];
          const textColor = readableTextColorFor(headerColor);
          return (
          <div key={group.name} className="overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm">
            <div className="flex items-center justify-between gap-2 px-3 py-1.5" style={{ background: headerColor }}>
              <span className="truncate text-xs font-bold uppercase tracking-wide" style={{ color: textColor }}>
                {group.name}
              </span>
              <span
                className="shrink-0 rounded-full px-2 py-0.5 text-[10px] font-bold"
                style={{ background: textColor === "#ffffff" ? "rgba(255,255,255,0.25)" : "rgba(31,36,48,0.12)", color: textColor }}
              >
                {group.fields.length}
              </span>
            </div>
            <ul>
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
          );
        })}
      </div>
    </SectionCard>
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
function ZoneDropArea({ zoneKey, onFieldDropped, isFieldAccepted, isEmpty, emptyText = "Drag a field here", children }) {
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
        isEmpty ? "min-h-[3rem] border-2 border-dashed border-slate-200 p-3" : ""
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
// Mode "Members" — the rule-based modes below never show a member list, they describe a
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
    if (!filter.field || mode !== "Members") {
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

          {mode === "Members" && (
            <MultiSelectField
              options={memberOptions}
              value={filter.includedMembers}
              onChange={(includedMembers) => onChange({ ...filter, includedMembers })}
              placeholder={loadingMembers ? "Loading members…" : "Type to search members…"}
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

// Excel's "Group Field" (date) and "Group Selection" (numeric bins), one control per Row field —
// backend enforces Rows-only and one grouping per field (see MdxPivotQueryBuilder.ValidateGroupings).
// Mode is derived from whichever of dateGroupings/numericGroupings currently mentions the field,
// rather than stored as its own separate piece of state, so the two arrays stay the single source
// of truth the backend actually reads.
function RowGroupingEditor({ rows, dateGroupings, numericGroupings, onChange }) {
  if (rows.length === 0) {
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
      <p className="text-xs font-medium text-muted">Group by (date/numeric fields only):</p>
      {rows.map((field) => {
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

function SortEditor({ sort, values, rows, columns, onChange }) {
  const mode = sort === null ? "none" : sort.customOrder ? "custom" : sort.byMeasureField ? sort.byMeasureField : "label";
  const axis = sort?.axis ?? "Rows";
  const targetField = (axis === "Columns" ? columns : rows)[0] ?? null;

  const handleModeChange = (nextMode) => {
    if (nextMode === "none") {
      onChange(null);
      return;
    }
    const direction = sort?.direction ?? "Ascending";
    if (nextMode === "custom") {
      onChange({ direction, byMeasureField: null, axis, customOrder: [] });
      return;
    }
    onChange(
      nextMode === "label"
        ? { direction, byMeasureField: null, axis, customOrder: null }
        : { direction, byMeasureField: nextMode, axis, customOrder: null },
    );
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
              onChange={(e) => onChange({ ...sort, axis: e.target.value })}
            >
              <option value="Rows">Sort Rows</option>
              <option value="Columns">Sort Columns</option>
            </select>
            {mode !== "custom" && (
              <select
                className={nativeSelectClass}
                value={sort.direction}
                onChange={(e) => onChange({ ...sort, direction: e.target.value })}
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
          onChange={(customOrder) => onChange({ ...sort, customOrder })}
        />
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
function ValuePill({ v, onUpdate, onUpdateFormat, onRemove, onDropReorder }) {
  const format = v.format ?? DEFAULT_FORMAT;

  return (
    <div onDragOver={(e) => e.preventDefault()} onDrop={onDropReorder}>
      <Popover
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

export default function FieldPicker({ value, onChange, dimensions, measures }) {
  const dimensionOptions = dimensions.map((d) => ({ value: d.field, label: d.displayName }));
  const measureOptions = measures.map((m) => ({ value: m.field, label: m.displayName }));
  // A calculated field's own LeftField/RightField must reference a real cube measure only (no
  // chaining calculated-on-calculated — see the backend's CalculatedField), but once declared its
  // name becomes usable anywhere a regular measure is, including as a Value.
  const calculatedFieldOptions = value.calculatedFields
    .filter((f) => f.name)
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
      <AvailableFieldsPanel dimensions={dimensions} isPlaced={isFieldPlaced} />

      {/* Compact "Drop Zones" strip — Filters/Columns/Rows/Values always visible as small boxes,
          each its own drop target + pill list. A pill shows only its field/measure name; the
          full add/edit form (dimension picker, filter mode, show-values-as, ...) lives in a
          Popover behind the box's "+" or, for Filters/Values, behind the pill itself — one click
          away instead of permanently occupying space, per-item since a filter/value pill needs
          its OWN config, not just a name. */}
      <div>
        <p className="mb-2 text-xs font-bold uppercase tracking-wide text-muted">Drop zones</p>
        <div className="grid grid-cols-1 items-stretch gap-3 sm:grid-cols-2 lg:grid-cols-4">
          <div className="flex h-full flex-col rounded-xl border border-slate-200 bg-white p-3 shadow-sm">
            <div className="mb-2 flex items-center justify-between gap-2">
              <span className="text-xs font-bold uppercase tracking-wide text-muted">
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

          <HierarchicalCubePathSelect
            label="Columns"
            options={dimensions.map((d) => d.field)}
            value={value.columns}
            onChange={(columns) =>
              onChange({ ...value, columns, calculatedItems: clearOrphanedCalculatedItems(value.calculatedItems, value.rows, columns) })
            }
            dragSourceKey="columns"
            onFieldDropped={moveField}
          />

          <HierarchicalCubePathSelect
            label="Rows"
            options={dimensions.map((d) => d.field)}
            value={value.rows}
            onChange={(rows) => {
              // Dropping a row field also drops any grouping/calculated item that referenced it —
              // the backend rejects both on a field that's no longer in Rows/Columns
              // (ValidateGroupings, ValidateCalculatedItems).
              const dateGroupings = value.dateGroupings.filter((g) => rows.includes(g.field));
              const numericGroupings = value.numericGroupings.filter((g) => rows.includes(g.field));
              const calculatedItems = clearOrphanedCalculatedItems(value.calculatedItems, rows, value.columns);
              onChange({ ...value, rows, dateGroupings, numericGroupings, calculatedItems });
            }}
            dragSourceKey="rows"
            onFieldDropped={moveField}
          />

          {/* No separate aggregation dropdown: confirmed against the real cube (2026-09-21) that
              measures are pre-built with their aggregation baked in — see PivotValueField in the
              backend for the full story. Measures are ADDED here via the multi-select (a measure
              has exactly one valid destination, so there's nothing for drag to decide — see
              AvailableFieldsPanel's WHY comment); a pill's own Popover then covers Show Values
              As/Format/Conditional format, and pills can still be drag-reordered among themselves
              since that order affects the generated MDX (see handleValueRowDrop). */}
          <div className="flex h-full flex-col rounded-xl border border-slate-200 bg-white p-3 shadow-sm">
            <div className="mb-2 flex items-center justify-between gap-2">
              <span className="text-xs font-bold uppercase tracking-wide text-muted">
                Values {value.values.length > 0 && <span className="text-gold">({value.values.length})</span>}
              </span>
            </div>
            <Select
              isMulti
              className="react-select-container mb-2"
              classNamePrefix="react-select"
              placeholder="Choose measures…"
              options={valueFieldOptions}
              value={valueFieldOptions.filter((opt) => value.values.some((v) => v.field === opt.value))}
              onChange={(selected) => setSelectedValueFields((selected ?? []).map((opt) => opt.value))}
              menuPortalTarget={typeof document !== "undefined" ? document.body : null}
              menuPosition="fixed"
              styles={{ menuPortal: (base) => ({ ...base, zIndex: 10000 }) }}
            />
            <select
              className={`${nativeSelectClass} mb-2 w-full`}
              value={value.valuesPlacement ?? "Columns"}
              onChange={(e) => onChange({ ...value, valuesPlacement: e.target.value })}
              title="Excel's own 'drag Σ Values between Columns and Rows' — see DECISIONS.md for this phase's scope cuts (needs a Column field, Normal-only Show Values As, no chart)"
            >
              <option value="Columns">Values on: Columns</option>
              <option value="Rows">Values on: Rows</option>
            </select>
            <ZoneDropArea
              zoneKey="values"
              onFieldDropped={moveField}
              isFieldAccepted={(field) => valueFieldOptions.some((opt) => opt.value === field)}
              isEmpty={value.values.length === 0}
              emptyText="Pick measures above"
            >
              <div className="flex flex-wrap gap-1.5">
                {value.values.map((v, index) => (
                  <ValuePill
                    key={index}
                    v={v}
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

        <RowGroupingEditor
          rows={value.rows}
          dateGroupings={value.dateGroupings}
          numericGroupings={value.numericGroupings}
          onChange={(dateGroupings, numericGroupings) => onChange({ ...value, dateGroupings, numericGroupings })}
        />
      </div>

      {/* Secondary, optional settings — small pill controls instead of full-width cards, each
          opening a Popover with its existing (unchanged) editor. */}
      <div className="flex flex-wrap items-center gap-2">
        <Popover
          trigger={(toggle) => (
            <button type="button" onClick={toggle} className="btn-secondary inline-flex items-center gap-1.5 px-3 py-1.5 text-xs">
              Sort
              {value.sort && <span className="rounded-full bg-gold/20 px-1.5 text-gold">1</span>}
              <ChevronDown className="h-3.5 w-3.5" />
            </button>
          )}
        >
          {() => (
            <SortEditor
              sort={value.sort}
              values={value.values}
              rows={value.rows}
              columns={value.columns}
              onChange={(sort) => onChange({ ...value, sort })}
            />
          )}
        </Popover>

        <Popover
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
              onChange={(calculatedFields) => onChange({ ...value, calculatedFields })}
            />
          )}
        </Popover>

        <Popover
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
