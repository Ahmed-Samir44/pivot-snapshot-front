import { useEffect, useState } from "react";
import Select from "react-select";
import { X } from "lucide-react";
import HierarchicalCubePathSelect from "../forms/HierarchicalCubePathSelect";
import MultiSelectField from "../forms/MultiSelectField";
import { getMembers } from "../../services/cubeMetaApi";

const SHOW_VALUES_AS_OPTIONS = [
  { value: "Normal", label: "Normal" },
  { value: "PercentOfGrandTotal", label: "% of Grand Total" },
  { value: "RunningTotal", label: "Running Total" },
  { value: "Rank", label: "Rank" },
];

const DEFAULT_FORMAT = { type: "General", decimalPlaces: 2, currencySymbol: "EGP" };

const nativeSelectClass =
  "rounded-xl border border-slate-200 bg-white/90 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent";

function SectionCard({ title, children }) {
  return (
    <div className="card">
      <h3 className="mb-4 text-sm font-bold uppercase tracking-wide text-muted">{title}</h3>
      {children}
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
    onChange({ field: filter.field, mode: nextMode, includedMembers: [], labelText: "", n: 10, byMeasureField: null });
  };

  return (
    <div className="mb-4 rounded-xl border border-slate-200 bg-slate-50/80 p-4">
      <div className="mb-3 flex items-start gap-3">
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
        </div>
      )}
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

  return (
    <div className="grid grid-cols-1 gap-6 lg:grid-cols-2">
      <SectionCard title="Rows">
        <HierarchicalCubePathSelect
          label="Row fields"
          options={dimensions.map((d) => d.field)}
          value={value.rows}
          onChange={(rows) => onChange({ ...value, rows })}
        />
      </SectionCard>

      <SectionCard title="Columns">
        <HierarchicalCubePathSelect
          label="Column fields"
          options={dimensions.map((d) => d.field)}
          value={value.columns}
          onChange={(columns) => onChange({ ...value, columns })}
        />
      </SectionCard>

      {/* Only a measure picker, no separate aggregation dropdown: confirmed against the real
          cube (2026-09-21) that measures are pre-built with their aggregation baked in — see
          PivotValueField in the backend for the full story. */}
      <SectionCard title="Values">
        {value.values.map((v, index) => {
          const format = v.format ?? DEFAULT_FORMAT;
          return (
            <div key={index} className="mb-3 rounded-xl border border-slate-100 p-3">
              <div className="flex flex-wrap items-center gap-2">
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
            </div>
          );
        })}
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

      <SectionCard title="Filters">
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
        <button type="button" onClick={addFilter} className="btn-secondary">
          + Add filter
        </button>
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
