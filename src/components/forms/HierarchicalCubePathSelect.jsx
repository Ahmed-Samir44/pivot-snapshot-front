// Ported from the Segmentation project's component of the same name — same two-step
// dimension -> field(s) picker, adapted to this project's cubeMeta utils. Replaces typing a raw
// MDX bracket path by hand with browsing the fields the backend's /api/cube/dimensions returns.
import { useMemo, useState } from "react";
import Select from "react-select";
import { Plus, X } from "lucide-react";
import { parseCubeBracketSegments, formatCubeBracketLabel, formatDimensionName } from "../../utils/cubeMeta";

const selectStyles = {
  menuPortal: (base) => ({ ...base, zIndex: 10000 }),
  control: (base) => ({
    ...base,
    minHeight: "48px",
    borderRadius: "12px",
    borderColor: "#e2e8f0",
    "&:hover": { borderColor: "#AE8C67" },
  }),
  multiValue: (base) => ({
    ...base,
    backgroundColor: "#AE8C67",
    borderRadius: "8px",
  }),
  multiValueLabel: (base) => ({
    ...base,
    color: "white",
    fontWeight: "600",
  }),
  multiValueRemove: (base) => ({
    ...base,
    color: "white",
    "&:hover": { backgroundColor: "#9a7b57" },
  }),
};

const singleSelectProps = {
  className: "react-select-container",
  classNamePrefix: "react-select",
  menuPortalTarget: typeof document !== "undefined" ? document.body : null,
  menuPosition: "fixed",
  styles: selectStyles,
  isClearable: true,
};

export default function HierarchicalCubePathSelect({
  options,
  value = [],
  onChange,
  label,
  placeholderAdd = "Choose dimension and field(s), then add — all matching paths are added at once.",
  selectedHeading = "Selected fields",
  addButtonLabel = "Add",
}) {
  const [step1, setStep1] = useState(null);
  const [step2Multi, setStep2Multi] = useState([]);

  const indexed = useMemo(
    () =>
      options
        .map((full) => ({ full, parts: parseCubeBracketSegments(full) }))
        .filter((x) => x.parts.length > 0),
    [options],
  );

  const opt1 = useMemo(() => {
    const set = new Set();
    indexed.forEach((row) => set.add(row.parts[0]));
    return [...set]
      .sort((a, b) => formatCubeBracketLabel(a).localeCompare(formatCubeBracketLabel(b)))
      .map((v) => ({ value: v, label: formatCubeBracketLabel(v) }));
  }, [indexed]);

  const after1 = useMemo(() => indexed.filter((row) => step1 && row.parts[0] === step1.value), [indexed, step1]);

  const opt2 = useMemo(() => {
    const set = new Set();
    after1.forEach((row) => {
      if (row.parts[1] != null) set.add(row.parts[1]);
    });
    return [...set]
      .sort((a, b) => formatCubeBracketLabel(a).localeCompare(formatCubeBracketLabel(b)))
      .map((v) => ({ value: v, label: formatCubeBracketLabel(v) }));
  }, [after1]);

  const selectedLevel2 = useMemo(() => new Set(step2Multi.map((o) => o.value)), [step2Multi]);
  const after2 = useMemo(() => after1.filter((row) => selectedLevel2.has(row.parts[1])), [after1, selectedLevel2]);

  const newPathsToAdd = useMemo(() => {
    if (!step1 || step2Multi.length === 0 || after2.length === 0) return [];
    const seen = new Set();
    const out = [];
    after2.forEach((row) => {
      if (seen.has(row.full) || value.includes(row.full)) return;
      seen.add(row.full);
      out.push(row.full);
    });
    return out;
  }, [step1, step2Multi, after2, value]);

  const canAdd = newPathsToAdd.length > 0;

  const handleAdd = () => {
    if (!canAdd) return;
    onChange([...value, ...newPathsToAdd]);
    setStep1(null);
    setStep2Multi([]);
  };

  const remove = (fullPath) => {
    onChange(value.filter((v) => v !== fullPath));
  };

  return (
    <div className="mb-6">
      <label className="mb-2 block text-sm font-semibold text-ink">{label}</label>

      <div className="space-y-4 rounded-xl border border-slate-200 bg-slate-50/80 p-4">
        <p className="text-sm text-muted">{placeholderAdd}</p>
        <div className="flex flex-wrap items-end gap-3">
          <div className="min-w-[min(100%,12rem)] flex-1 basis-[10rem]">
            <span className="mb-1 block text-xs font-medium text-muted">1 · Dimension</span>
            <Select
              {...singleSelectProps}
              placeholder="Select dimension…"
              options={opt1}
              value={step1}
              onChange={(v) => {
                setStep1(v);
                setStep2Multi([]);
              }}
            />
          </div>
          <div className="min-w-[min(100%,12rem)] flex-1 basis-[10rem]">
            <span className="mb-1 block text-xs font-medium text-muted">2 · Field(s)</span>
            <Select
              isMulti
              className="react-select-container"
              classNamePrefix="react-select"
              menuPortalTarget={typeof document !== "undefined" ? document.body : null}
              menuPosition="fixed"
              styles={selectStyles}
              isClearable
              closeMenuOnSelect={false}
              hideSelectedOptions={false}
              placeholder={step1 ? "Select field(s)…" : "Choose dimension first"}
              options={opt2}
              value={step2Multi}
              isDisabled={!step1}
              onChange={(v) => setStep2Multi(v || [])}
            />
          </div>
          <button
            type="button"
            onClick={handleAdd}
            disabled={!canAdd}
            className="btn-primary inline-flex h-12 shrink-0 items-center justify-center gap-2 whitespace-nowrap px-4 disabled:cursor-not-allowed disabled:opacity-50"
          >
            <Plus className="h-5 w-5 shrink-0" />
            {addButtonLabel}
          </button>
        </div>
      </div>

      {value.length > 0 && (
        <div className="mt-4">
          <p className="mb-2 text-xs font-medium text-muted">
            {selectedHeading} ({value.length})
          </p>
          <ul className="flex flex-wrap gap-2">
            {value.map((fullPath) => (
              <li
                key={fullPath}
                className="inline-flex items-center gap-1.5 rounded-lg bg-gold px-3 py-1.5 text-sm font-medium text-white"
              >
                <span>{formatDimensionName(fullPath)}</span>
                <button
                  type="button"
                  onClick={() => remove(fullPath)}
                  className="rounded p-0.5 hover:bg-white/20"
                  aria-label="Remove"
                >
                  <X className="h-4 w-4" />
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
}
