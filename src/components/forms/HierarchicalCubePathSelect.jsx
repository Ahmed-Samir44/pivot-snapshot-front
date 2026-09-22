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

// Native HTML5 drag-and-drop (not a library) — dataTransfer carries a plain string across the
// browser's DOM event system, which works between two SEPARATE instances of this component
// (Rows and Columns each render their own) without any shared React state between them. The
// actual cross-list move (remove from one array, add to the other) still has to happen one level
// up, in FieldPicker, since only it holds both arrays — see onFieldDropped.
const DRAG_MIME_TYPE = "application/x-pivot-field";

export default function HierarchicalCubePathSelect({
  options,
  value = [],
  onChange,
  label,
  placeholderAdd = "Choose dimension and field(s), then add — all matching paths are added at once.",
  selectedHeading = "Selected fields",
  addButtonLabel = "Add",
  dragSourceKey = null,
  onFieldDropped = null,
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

  const [dragOver, setDragOver] = useState(false);

  const handleDragStart = (e, fullPath) => {
    if (!dragSourceKey) return;
    e.dataTransfer.setData(DRAG_MIME_TYPE, JSON.stringify({ field: fullPath, source: dragSourceKey }));
    e.dataTransfer.effectAllowed = "move";
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    setDragOver(true);
  };

  const handleDrop = (e) => {
    if (!dragSourceKey || !onFieldDropped) return;
    e.preventDefault();
    setDragOver(false);
    const raw = e.dataTransfer.getData(DRAG_MIME_TYPE);
    if (!raw) return;
    const { field, source } = JSON.parse(raw);
    if (source !== dragSourceKey) {
      onFieldDropped(field, source, dragSourceKey);
    }
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

      {(value.length > 0 || (dragSourceKey && onFieldDropped)) && (
        <div className="mt-4">
          <p className="mb-2 text-xs font-medium text-muted">
            {selectedHeading} ({value.length})
            {dragSourceKey && onFieldDropped && " — drag a field here to move it from the other axis"}
          </p>
          <ul
            className={`flex flex-wrap gap-2 rounded-lg ${dragOver ? "bg-gold/10 ring-2 ring-gold ring-inset" : ""} ${
              value.length === 0 ? "min-h-[2.5rem] border-2 border-dashed border-slate-200 p-2" : ""
            }`}
            onDragOver={dragSourceKey && onFieldDropped ? handleDragOver : undefined}
            onDragLeave={dragSourceKey && onFieldDropped ? () => setDragOver(false) : undefined}
            onDrop={dragSourceKey && onFieldDropped ? handleDrop : undefined}
          >
            {value.map((fullPath) => (
              <li
                key={fullPath}
                draggable={Boolean(dragSourceKey)}
                onDragStart={(e) => handleDragStart(e, fullPath)}
                className={`inline-flex items-center gap-1.5 rounded-lg bg-gold px-3 py-1.5 text-sm font-medium text-white ${
                  dragSourceKey ? "cursor-grab active:cursor-grabbing" : ""
                }`}
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
