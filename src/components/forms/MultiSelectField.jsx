// Ported from the Segmentation project — a plain multi-select styled to match the rest of the
// gold/ink design system. Used here for picking a Filter's included members.
import Select from "react-select";

export default function MultiSelectField({ options, value, onChange, label, placeholder = "Select options...", isLoading = false }) {
  const formattedOptions = options.map((opt) => ({ value: opt, label: opt }));
  const selectedValues = Array.isArray(value)
    ? value.map((v) => formattedOptions.find((opt) => opt.value === v) || { value: v, label: v })
    : [];

  return (
    <div className="mb-2">
      {label && <label className="mb-2 block text-sm font-semibold text-ink">{label}</label>}
      <Select
        isMulti
        isLoading={isLoading}
        value={selectedValues}
        onChange={(selected) => onChange(selected ? selected.map((opt) => opt.value) : [])}
        options={formattedOptions}
        placeholder={placeholder}
        isClearable
        isSearchable
        menuPortalTarget={typeof document !== "undefined" ? document.body : null}
        menuPosition="fixed"
        className="react-select-container"
        classNamePrefix="react-select"
        styles={{
          menuPortal: (base) => ({ ...base, zIndex: 10000 }),
          control: (base) => ({
            ...base,
            minHeight: "48px",
            borderRadius: "12px",
            borderColor: "#e2e8f0",
            "&:hover": { borderColor: "#AE8C67" },
          }),
          multiValue: (base) => ({ ...base, backgroundColor: "#AE8C67", borderRadius: "8px" }),
          multiValueLabel: (base) => ({ ...base, color: "white", fontWeight: "600" }),
          multiValueRemove: (base) => ({ ...base, color: "white", "&:hover": { backgroundColor: "#9a7b57" } }),
        }}
      />
    </div>
  );
}
