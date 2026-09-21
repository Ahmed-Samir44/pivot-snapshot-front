// Rows/Columns/Values/Filters as free-text field identifiers (e.g. "[Doctor].[Doctor Name]"),
// plus Sort and Show Grand Totals controls. No drag-and-drop or member picker yet — those are
// later Excel-parity work, see DECISIONS.md.
const AGGREGATIONS = ["Sum", "Count", "Average", "Min", "Max"];
const SHOW_VALUES_AS_OPTIONS = [
  { value: "Normal", label: "عادي" },
  { value: "PercentOfGrandTotal", label: "% من الإجمالي الكلي" },
];

function FieldListEditor({ label, placeholder, fields, onChange }) {
  const updateField = (index, next) => {
    const copy = [...fields];
    copy[index] = next;
    onChange(copy);
  };

  const removeField = (index) => {
    onChange(fields.filter((_, i) => i !== index));
  };

  return (
    <fieldset className="field-editor">
      <legend>{label}</legend>
      {fields.map((field, index) => (
        <div className="field-row" key={index}>
          <input
            type="text"
            value={field}
            placeholder={placeholder}
            onChange={(e) => updateField(index, e.target.value)}
          />
          <button type="button" onClick={() => removeField(index)} aria-label={`إزالة ${label}`}>
            ✕
          </button>
        </div>
      ))}
      <button type="button" onClick={() => onChange([...fields, ""])}>
        + إضافة {label}
      </button>
    </fieldset>
  );
}

// A filter selects specific members of a dimension (comma-separated) — matches Excel's Filters
// area, or a member-level filter on a field also used as a Row/Column (see MdxPivotQueryBuilder).
function FilterListEditor({ filters, onChange }) {
  const updateFilter = (index, patch) => {
    const copy = [...filters];
    copy[index] = { ...copy[index], ...patch };
    onChange(copy);
  };

  const removeFilter = (index) => {
    onChange(filters.filter((_, i) => i !== index));
  };

  const addFilter = () => {
    onChange([...filters, { field: "", membersText: "" }]);
  };

  return (
    <fieldset className="field-editor">
      <legend>Filters</legend>
      {filters.map((filter, index) => (
        <div className="field-row" key={index}>
          <input
            type="text"
            value={filter.field}
            placeholder="[Date].[Year]"
            onChange={(e) => updateFilter(index, { field: e.target.value })}
          />
          <input
            type="text"
            value={filter.membersText}
            placeholder="2025, 2026"
            onChange={(e) => updateFilter(index, { membersText: e.target.value })}
          />
          <button type="button" onClick={() => removeFilter(index)} aria-label="إزالة Filter">
            ✕
          </button>
        </div>
      ))}
      <button type="button" onClick={addFilter}>
        + إضافة Filter
      </button>
    </fieldset>
  );
}

// Sort applies to the ROWS axis only (see MdxPivotQueryBuilder). "By label" is only valid when
// there's exactly one row field — the backend rejects it otherwise, so the dropdown here doesn't
// try to pre-validate that; the API's error message explains it if the user hits that case.
function SortEditor({ sort, values, onChange }) {
  const mode = sort === null ? "none" : sort.byMeasureField ? sort.byMeasureField : "label";

  const handleModeChange = (nextMode) => {
    if (nextMode === "none") {
      onChange(null);
      return;
    }
    const direction = sort?.direction ?? "Ascending";
    onChange(nextMode === "label" ? { direction, byMeasureField: null } : { direction, byMeasureField: nextMode });
  };

  return (
    <fieldset className="field-editor">
      <legend>Sort</legend>
      <div className="field-row">
        <select value={mode} onChange={(e) => handleModeChange(e.target.value)}>
          <option value="none">بدون ترتيب</option>
          <option value="label">بالاسم (Row واحد بس)</option>
          {values.filter((v) => v.field).map((v) => (
            <option key={v.field} value={v.field}>
              بقيمة {v.field}
            </option>
          ))}
        </select>
        {sort && (
          <select
            value={sort.direction}
            onChange={(e) => onChange({ ...sort, direction: e.target.value })}
          >
            <option value="Ascending">تصاعدي</option>
            <option value="Descending">تنازلي</option>
          </select>
        )}
      </div>
    </fieldset>
  );
}

export default function FieldPicker({ value, onChange }) {
  const updateValue = (index, patch) => {
    const copy = [...value.values];
    copy[index] = { ...copy[index], ...patch };
    onChange({ ...value, values: copy });
  };

  const removeValue = (index) => {
    onChange({ ...value, values: value.values.filter((_, i) => i !== index) });
  };

  const addValue = () => {
    onChange({ ...value, values: [...value.values, { field: "", aggregation: "Sum", showValuesAs: "Normal" }] });
  };

  return (
    <div className="field-picker">
      <FieldListEditor
        label="Rows"
        placeholder="[Doctor].[Doctor Name]"
        fields={value.rows}
        onChange={(rows) => onChange({ ...value, rows })}
      />
      <FieldListEditor
        label="Columns"
        placeholder="[Date].[Month]"
        fields={value.columns}
        onChange={(columns) => onChange({ ...value, columns })}
      />
      <fieldset className="field-editor">
        <legend>Values</legend>
        {value.values.map((v, index) => (
          <div className="field-row" key={index}>
            <input
              type="text"
              value={v.field}
              placeholder="VisitCount"
              onChange={(e) => updateValue(index, { field: e.target.value })}
            />
            <select
              value={v.aggregation}
              onChange={(e) => updateValue(index, { aggregation: e.target.value })}
            >
              {AGGREGATIONS.map((agg) => (
                <option key={agg} value={agg}>
                  {agg}
                </option>
              ))}
            </select>
            <select
              value={v.showValuesAs}
              onChange={(e) => updateValue(index, { showValuesAs: e.target.value })}
            >
              {SHOW_VALUES_AS_OPTIONS.map((opt) => (
                <option key={opt.value} value={opt.value}>
                  {opt.label}
                </option>
              ))}
            </select>
            <button type="button" onClick={() => removeValue(index)} aria-label="إزالة Value">
              ✕
            </button>
          </div>
        ))}
        <button type="button" onClick={addValue}>
          + إضافة Value
        </button>
      </fieldset>
      <FilterListEditor filters={value.filters} onChange={(filters) => onChange({ ...value, filters })} />
      <SortEditor sort={value.sort} values={value.values} onChange={(sort) => onChange({ ...value, sort })} />
      <fieldset className="field-editor">
        <legend>Totals</legend>
        <label className="field-row">
          <input
            type="checkbox"
            checked={value.showGrandTotals}
            onChange={(e) => onChange({ ...value, showGrandTotals: e.target.checked })}
          />
          Show Grand Totals
        </label>
      </fieldset>
    </div>
  );
}
