import { useEffect, useState } from "react";
import Select from "react-select";
import { X, GripVertical } from "lucide-react";
import HierarchicalCubePathSelect from "../forms/HierarchicalCubePathSelect";
import MultiSelectField from "../forms/MultiSelectField";
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

// Excel's Field List: every cube field in one place, dragged from here into whichever of the four
// zones below it belongs in. Unlike the zones, a field never "leaves" this list when placed
// somewhere — Excel's own field list keeps every field checkbox visible too — it just gets a
// small dot marking it as already used somewhere, since (unlike Excel) this tool allows the same
// field to sit in Filters AND Rows/Columns/Values at once (a Row/Column field can independently
// carry its own member filter — see moveField's WHY comment in FieldPicker).
function AvailableFieldsPanel({ dimensions, measures, calculatedFieldNames, isPlaced }) {
  const allFields = [
    ...dimensions.map((d) => ({ field: d.field, label: d.displayName })),
    ...measures.map((m) => ({ field: m.field, label: m.displayName })),
    ...calculatedFieldNames.map((name) => ({ field: name, label: `${name} (calculated)` })),
  ];

  return (
    <SectionCard title="Available fields">
      <p className="mb-3 text-xs text-muted">Drag a field into Filters, Rows, Columns, or Values below.</p>
      <ul className="flex max-h-64 flex-wrap gap-2 overflow-y-auto">
        {allFields.map(({ field, label }) => (
          <li
            key={field}
            draggable
            onDragStart={(e) => setFieldDragPayload(e, field, "available")}
            className="inline-flex cursor-grab items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-sm font-medium text-ink shadow-sm active:cursor-grabbing"
          >
            {isPlaced(field) && <span className="h-1.5 w-1.5 rounded-full bg-gold" aria-hidden="true" />}
            {label}
          </li>
        ))}
      </ul>
    </SectionCard>
  );
}

// A generic drop target for the Values and Filters zones (Rows/Columns use
// HierarchicalCubePathSelect's own built-in drop handling instead — see that component). Visually
// matches the dashed-border "drop here" affordance in HierarchicalCubePathSelect for consistency
// across all four zones.
function ZoneDropArea({ zoneKey, onFieldDropped, isEmpty, children }) {
  const [dragOver, setDragOver] = useState(false);

  const handleDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    const payload = readFieldDragPayload(e);
    if (payload && payload.source !== zoneKey) {
      onFieldDropped(payload.field, payload.source, zoneKey);
    }
  };

  return (
    <div
      className={`rounded-lg ${dragOver ? "bg-gold/10 ring-2 ring-gold ring-inset" : ""} ${
        isEmpty ? "min-h-[3rem] border-2 border-dashed border-slate-200 p-3" : ""
      }`}
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={handleDrop}
    >
      {isEmpty ? <p className="text-xs text-muted">Drag a field here</p> : children}
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
function FilterRow({ filter, dimensionOptions, selectedValueOptions, onChange, onRemove }) {
  const [memberOptions, setMemberOptions] = useState([]);
  const [loadingMembers, setLoadingMembers] = useState(false);
  const mode = filter.mode ?? "Members";

  useEffect(() => {
    if (!filter.field || mode !== "Members") {
      setMemberOptions([]);
      return;
    }
    let cancelled = false;
    setLoadingMembers(true);
    getMembers(filter.field)
      .then((members) => {
        if (!cancelled) setMemberOptions(members);
      })
      .finally(() => {
        if (!cancelled) setLoadingMembers(false);
      });
    return () => {
      cancelled = true;
    };
  }, [filter.field, mode]);

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
              placeholder={loadingMembers ? "Loading members…" : "Select member(s) to include…"}
              isLoading={loadingMembers}
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

  const addValue = () => {
    onChange({ ...value, values: [...value.values, { field: "", showValuesAs: "Normal" }] });
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
    <div className="grid grid-cols-1 gap-6 lg:grid-cols-2">
      <div className="lg:col-span-2">
        <AvailableFieldsPanel
          dimensions={dimensions}
          measures={measures}
          calculatedFieldNames={value.calculatedFields.filter((f) => f.name).map((f) => f.name)}
          isPlaced={isFieldPlaced}
        />
      </div>

      {/* Layout order matches Excel's own Field List pane: Filters + Columns on top,
          Rows + Values below. */}
      <SectionCard title="Filters" badge={value.filters.length}>
        <ZoneDropArea zoneKey="filters" onFieldDropped={moveField} isEmpty={value.filters.length === 0}>
          {value.filters.map((filter, index) => (
            <FilterRow
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
        </ZoneDropArea>
        <button type="button" onClick={addFilter} className="btn-secondary">
          + Add filter
        </button>
      </SectionCard>

      <SectionCard title="Columns" badge={value.columns.length}>
        <HierarchicalCubePathSelect
          label="Column fields"
          options={dimensions.map((d) => d.field)}
          value={value.columns}
          onChange={(columns) => onChange({ ...value, columns })}
          dragSourceKey="columns"
          onFieldDropped={moveField}
        />
      </SectionCard>

      <SectionCard title="Rows" badge={value.rows.length}>
        <HierarchicalCubePathSelect
          label="Row fields"
          options={dimensions.map((d) => d.field)}
          value={value.rows}
          onChange={(rows) => {
            // Dropping a row field also drops any grouping that referenced it — the backend
            // rejects a grouping on a field that's no longer in Rows (ValidateGroupings).
            const dateGroupings = value.dateGroupings.filter((g) => rows.includes(g.field));
            const numericGroupings = value.numericGroupings.filter((g) => rows.includes(g.field));
            onChange({ ...value, rows, dateGroupings, numericGroupings });
          }}
          dragSourceKey="rows"
          onFieldDropped={moveField}
        />
        <RowGroupingEditor
          rows={value.rows}
          dateGroupings={value.dateGroupings}
          numericGroupings={value.numericGroupings}
          onChange={(dateGroupings, numericGroupings) => onChange({ ...value, dateGroupings, numericGroupings })}
        />
      </SectionCard>

      {/* Only a measure picker, no separate aggregation dropdown: confirmed against the real
          cube (2026-09-21) that measures are pre-built with their aggregation baked in — see
          PivotValueField in the backend for the full story. */}
      <SectionCard title="Values" badge={value.values.length}>
        <ZoneDropArea zoneKey="values" onFieldDropped={moveField} isEmpty={value.values.length === 0}>
        {value.values.map((v, index) => {
          const format = v.format ?? DEFAULT_FORMAT;
          return (
            <div
              key={index}
              className="mb-3 rounded-xl border border-slate-100 p-3"
              onDragOver={v.field ? (e) => e.preventDefault() : undefined}
              onDrop={v.field ? (e) => handleValueRowDrop(e, index) : undefined}
            >
              <div className="flex flex-wrap items-center gap-2">
                {v.field && (
                  <span
                    draggable
                    onDragStart={(e) => setFieldDragPayload(e, v.field, "values")}
                    className="cursor-grab text-muted hover:text-ink active:cursor-grabbing"
                    title="Drag to move this field to another zone, or drop onto another value row to reorder"
                  >
                    <GripVertical className="h-5 w-5" />
                  </span>
                )}
                <div className="min-w-[10rem] flex-1">
                  <Select
                    className="react-select-container"
                    classNamePrefix="react-select"
                    placeholder="Choose a measure…"
                    options={valueFieldOptions}
                    value={valueFieldOptions.find((opt) => opt.value === v.field) ?? null}
                    onChange={(opt) => updateValue(index, { field: opt?.value ?? "" })}
                    menuPortalTarget={typeof document !== "undefined" ? document.body : null}
                    menuPosition="fixed"
                    styles={{ menuPortal: (base) => ({ ...base, zIndex: 10000 }) }}
                  />
                </div>
                <select
                  className={nativeSelectClass}
                  value={v.showValuesAs}
                  onChange={(e) => updateValue(index, { showValuesAs: e.target.value })}
                >
                  {SHOW_VALUES_AS_OPTIONS.map((opt) => (
                    <option key={opt.value} value={opt.value}>
                      {opt.label}
                    </option>
                  ))}
                </select>
                <button type="button" onClick={() => removeValue(index)} className="text-muted hover:text-ink" aria-label="Remove value">
                  <X className="h-5 w-5" />
                </button>
              </div>
              <div className="mt-2 flex flex-wrap items-center gap-2">
                <span className="text-xs font-medium text-muted">Format:</span>
                <select
                  className={nativeSelectClass}
                  value={format.type}
                  onChange={(e) => updateValueFormat(index, { type: e.target.value })}
                >
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
                    onChange={(e) => updateValueFormat(index, { decimalPlaces: Number(e.target.value) })}
                    className="input-field w-20 py-2"
                    title="Decimal places"
                  />
                )}
                {format.type === "Currency" && (
                  <input
                    type="text"
                    value={format.currencySymbol}
                    onChange={(e) => updateValueFormat(index, { currencySymbol: e.target.value })}
                    placeholder="EGP"
                    className="input-field w-24 py-2"
                    title="Currency symbol"
                  />
                )}
              </div>
              <div className="mt-2 flex flex-wrap items-center gap-2">
                <span className="text-xs font-medium text-muted">Conditional format:</span>
                <select
                  className={nativeSelectClass}
                  value={v.conditionalFormat?.type ?? "None"}
                  onChange={(e) => updateValue(index, { conditionalFormat: { type: e.target.value } })}
                >
                  <option value="None">None</option>
                  <option value="ColorScale">Color scale</option>
                  <option value="DataBar">Data bar</option>
                  <option value="IconSet">Icon set (traffic lights)</option>
                </select>
              </div>
            </div>
          );
        })}
        </ZoneDropArea>
        <button type="button" onClick={addValue} className="btn-secondary mt-1">
          + Add value
        </button>
      </SectionCard>

      <SectionCard title="Calculated Fields">
        <CalculatedFieldsEditor
          calculatedFields={value.calculatedFields}
          measureOptions={measureOptions}
          onChange={(calculatedFields) => onChange({ ...value, calculatedFields })}
        />
      </SectionCard>

      <SectionCard title="Sort">
        <SortEditor
          sort={value.sort}
          values={value.values}
          rows={value.rows}
          columns={value.columns}
          onChange={(sort) => onChange({ ...value, sort })}
        />
      </SectionCard>

      <SectionCard title="Totals">
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
      </SectionCard>
    </div>
  );
}
