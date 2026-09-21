// Rows/Columns/Values only, as free-text field identifiers (e.g. "[Doctor].[Doctor Name]").
// No Filters area, drag-and-drop, or member picker yet — those are Phase B (Excel parity) work,
// see DECISIONS.md.
const AGGREGATIONS = ["Sum", "Count", "Average", "Min", "Max"];

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
    onChange({ ...value, values: [...value.values, { field: "", aggregation: "Sum" }] });
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
            <button type="button" onClick={() => removeValue(index)} aria-label="إزالة Value">
              ✕
            </button>
          </div>
        ))}
        <button type="button" onClick={addValue}>
          + إضافة Value
        </button>
      </fieldset>
    </div>
  );
}
