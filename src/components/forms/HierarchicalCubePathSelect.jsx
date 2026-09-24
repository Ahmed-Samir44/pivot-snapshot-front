// Ported from the Segmentation project's component of the same name — same two-step
// dimension -> field(s) picker, adapted to this project's cubeMeta utils. Replaces typing a raw
// MDX bracket path by hand with browsing the fields the backend's /api/cube/dimensions returns.
import { useMemo, useState } from "react";
import Select from "react-select";
import { Plus, X } from "lucide-react";
import Popover from "./Popover";
import { parseCubeBracketSegments, formatCubeBracketLabel, formatDimensionName } from "../../utils/cubeMeta";
import { setFieldDragPayload, readFieldDragPayload } from "../../utils/dragDrop";

const selectStyles = {
  menuPortal: (base) => ({ ...base, zIndex: 10000 }),
  control: (base) => ({
    ...base,
    borderRadius: "10px",
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

// Native HTML5 drag-and-drop (not a library, see utils/dragDrop.js) — works between this
// component's own two instances (Rows and Columns) AND the Values/Filters zones and the
// Available Fields panel elsewhere in FieldPicker, none of which share React state with this
// component. The actual move (remove from one zone's array, add to another's, converting shape
// where needed) happens one level up in FieldPicker, since only it holds every zone — see
// onFieldDropped.
//
// Compact drop-zone box: the pill list (and its drag/drop target) is always visible, but the
// two-step Dimension→Field(s) picker only exists inside a Popover behind the "+" button — moved
// there so the zone reads as a small box (matching the rest of FieldPicker's compact "Drop
// Zones" strip) instead of an always-expanded form taking the same space whether or not the user
// is actively adding a field right now.
export default function HierarchicalCubePathSelect({
  options,
  value = [],
  onChange,
  label,
  placeholderAdd = "Choose dimension and field(s), then add — all matching paths are added at once.",
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

  const handleAdd = (closePopover) => {
    if (!canAdd) return;
    onChange([...value, ...newPathsToAdd]);
    setStep1(null);
    setStep2Multi([]);
    closePopover();
  };

  const remove = (fullPath) => {
    onChange(value.filter((v) => v !== fullPath));
  };

  const [dragOver, setDragOver] = useState(false);

  const handleDragStart = (e, fullPath) => {
    if (!dragSourceKey) return;
    setFieldDragPayload(e, fullPath, dragSourceKey);
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    setDragOver(true);
  };

  const handleDrop = (e) => {
    if (!dragSourceKey || !onFieldDropped) return;
    e.preventDefault();
    setDragOver(false);
    const payload = readFieldDragPayload(e);
    if (payload && payload.source !== dragSourceKey) {
      onFieldDropped(payload.field, payload.source, dragSourceKey);
    }
  };

  // Dropping directly ON a pill (rather than empty space in the list) reorders WITHIN this same
  // zone when the dragged field came from here too — stopPropagation keeps the container's own
  // onDrop (cross-zone move) from also firing for the same event. A drop from another zone landing
  // on a pill still falls through to the ordinary cross-zone move (appended at the end), same as
  // dropping on empty space; only same-zone drags get positional reordering here.
  const handlePillDrop = (e, targetFullPath) => {
    if (!dragSourceKey) return;
    e.preventDefault();
    e.stopPropagation();
    setDragOver(false);
    const payload = readFieldDragPayload(e);
    if (!payload) return;
    if (payload.source === dragSourceKey) {
      if (payload.field === targetFullPath) return;
      const withoutDragged = value.filter((v) => v !== payload.field);
      const targetIndex = withoutDragged.indexOf(targetFullPath);
      if (targetIndex === -1) return;
      onChange([...withoutDragged.slice(0, targetIndex), payload.field, ...withoutDragged.slice(targetIndex)]);
    } else if (onFieldDropped) {
      onFieldDropped(payload.field, payload.source, dragSourceKey);
    }
  };

  return (
    // flex h-full flex-col: this box sits in a CSS grid row (FieldPicker's "Drop zones" strip)
    // that stretches every card to match the tallest sibling — without this, the <ul> drop target
    // below stayed sized to its own pill content, leaving genuine leftover white space inside a
    // stretched card that LOOKED droppable but wasn't (caught live, 2026-09-23). flex-1 on the
    // <ul> below is what actually claims that leftover space.
    <div className="flex h-full flex-col rounded-xl border border-slate-200 bg-white p-3 shadow-sm">
      <div className="mb-2 flex items-center justify-between gap-2">
        <span className="text-xs font-bold uppercase tracking-wide text-muted">
          {label} {value.length > 0 && <span className="text-gold">({value.length})</span>}
        </span>
        <Popover
          trigger={(toggle) => (
            <button
              type="button"
              onClick={toggle}
              className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-gold/15 text-gold hover:bg-gold/25"
              aria-label={`Add a field to ${label}`}
              title={`Add a field to ${label}`}
            >
              <Plus className="h-3.5 w-3.5" />
            </button>
          )}
        >
          {(close) => (
            <div className="space-y-3">
              <p className="text-xs text-muted">{placeholderAdd}</p>
              <div>
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
              <div>
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
                onClick={() => handleAdd(close)}
                disabled={!canAdd}
                className="btn-primary inline-flex w-full items-center justify-center gap-1.5 whitespace-nowrap px-4 py-2 disabled:cursor-not-allowed disabled:opacity-50"
              >
                <Plus className="h-4 w-4 shrink-0" />
                {addButtonLabel}
              </button>
            </div>
          )}
        </Popover>
      </div>

      {/* items-start + content-start: the <ul> itself is the flex-1 element being stretched to
          fill the card's leftover height (see the WHY comment above), but it's ALSO a flex
          container for its own pill <li>s — flex's default align-items/align-content is
          "stretch", which was blowing each pill up to fill that whole leftover height instead of
          leaving it empty below a normal-sized pill (caught live, 2026-09-23). Pinning both to
          "start" keeps the pills their natural size while the extra height still counts as part
          of the droppable area (the <ul> itself still gets the drop handlers). */}
      <ul
        className={`flex min-h-[2.25rem] flex-1 flex-wrap content-start gap-1.5 rounded-lg ${value.length === 0 ? "items-center" : "items-start"} ${dragOver ? "bg-gold/10 ring-2 ring-gold ring-inset" : ""} ${
          value.length === 0 ? "border-2 border-dashed border-slate-200 px-2" : ""
        }`}
        onDragOver={dragSourceKey && onFieldDropped ? handleDragOver : undefined}
        onDragLeave={dragSourceKey && onFieldDropped ? () => setDragOver(false) : undefined}
        onDrop={dragSourceKey && onFieldDropped ? handleDrop : undefined}
      >
        {value.length === 0 && <span className="text-xs text-muted">Drop a field here</span>}
        {value.map((fullPath) => (
          <li
            key={fullPath}
            draggable={Boolean(dragSourceKey)}
            onDragStart={(e) => handleDragStart(e, fullPath)}
            onDragOver={dragSourceKey ? (e) => e.preventDefault() : undefined}
            onDrop={dragSourceKey ? (e) => handlePillDrop(e, fullPath) : undefined}
            className={`inline-flex items-center gap-1 rounded-md bg-gold px-2 py-1 text-xs font-medium text-white ${
              dragSourceKey ? "cursor-grab active:cursor-grabbing" : ""
            }`}
          >
            <span>{formatDimensionName(fullPath)}</span>
            <button type="button" onClick={() => remove(fullPath)} className="rounded p-0.5 hover:bg-white/20" aria-label="Remove">
              <X className="h-3 w-3" />
            </button>
          </li>
        ))}
      </ul>
    </div>
  );
}
