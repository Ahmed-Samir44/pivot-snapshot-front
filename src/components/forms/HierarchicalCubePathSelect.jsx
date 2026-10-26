// Ported from the Segmentation project's component of the same name — same two-step
// dimension -> field(s) picker, adapted to this project's cubeMeta utils. Replaces typing a raw
// MDX bracket path by hand with browsing the fields the backend's /api/cube/dimensions returns.
import { useMemo, useState } from "react";
import Select from "react-select";
import { Plus, ChevronDown } from "lucide-react";
import Popover from "./Popover";
import { parseCubeBracketSegments, formatCubeBracketLabel, formatDimensionName, fieldCaption } from "../../utils/cubeMeta";
import { setFieldDragPayload, readFieldDragPayload } from "../../utils/dragDrop";
import { useFieldDragActive } from "../../utils/useFieldDragActive";
import { MEASURES_PLACEHOLDER } from "../../utils/measuresPlacement";
import { effectiveSubtotalPosition } from "../../utils/subtotals";
import { SUMMARY_FUNCTIONS, aggregationLabel } from "../../utils/aggregations";

// One placed field pill (Rows or Columns) — its own component, not inlined in the .map() below,
// for the same reason it always was (hooks can't live inside a loop whose length changes).
function PlacedFieldPill({
  fullPath,
  dragSourceKey,
  onDragStart,
  onPillDrop,
  onRemove,
  contextMenuTargets,
  onMoveTo,
  isFirst,
  isLast,
  onMoveUp,
  onMoveDown,
  onMoveToBeginning,
  onMoveToEnd,
  // Real Excel's per-field "Field Settings…" (2026-10-13) — only passed for Rows (Subtotals is a
  // Rows-only feature, see PivotBuilder.jsx's own call site); undefined here means "don't show this
  // menu item at all" rather than showing a non-functional one, so a Columns pill's menu stays
  // exactly as it was. { reportDefault, overrides, onChange(field, patch) } — see FieldPicker.jsx.
  fieldSettings,
  // Custom Name / Layout & Print half of the same dialog (PLAN 6 Phase 2) — a SEPARATE prop from
  // fieldSettings because it's backed by different state one level up (see FieldPicker.jsx's own
  // WHY comment on its fieldLayout prop). { overrides, onChange(field, patch) }.
  fieldLayout,
  // Show Items With No Data, per field (PLAN 7, EXPERIMENTAL) — a THIRD prop, back in
  // value/onChange like fieldSettings (query-affecting), not the sibling fieldLayout. { fields,
  // onChange(field, checked) }.
  fieldShowNoData,
}) {
  // Columns (2026-10-14) passes fieldSettings for its drilldown field only — appliesTo(field) picks
  // it, and noCustom hides "Custom" (column totals have no function choice here).
  if (fieldSettings?.appliesTo && !fieldSettings.appliesTo(fullPath)) {
    fieldSettings = undefined;
  }
  // "menu" (the root Move/Remove list) or "subtotals" (the Field Settings sub-panel) — reset to
  // "menu" from the trigger button's own onClick below (same pattern FieldHeaderMenu's own
  // setPanel("shortcuts") already uses in PivotGrid.jsx), not on unmount, since this component
  // itself doesn't unmount/remount across a close/reopen the way the Popover's PANEL content does.
  const [panel, setPanel] = useState("menu");
  // Σ Values can move to the OTHER dimension axis (Rows<->Columns) but never to Filters or Values
  // itself — it isn't a real field, so both are dropped for it specifically, unlike every other
  // target.
  const isMeasuresPlaceholder = fullPath === MEASURES_PLACEHOLDER;
  const moveTargets = isMeasuresPlaceholder
    ? contextMenuTargets.filter((t) => t.zoneKey !== "filters" && t.zoneKey !== "values")
    : contextMenuTargets;
  // Menu shape and order match a real Excel field pill's own dropdown exactly (confirmed live
  // 2026-10-13 against a screenshot of it): Move Up/Down/Beginning/End first (grayed out at
  // whichever ends don't apply, same as Excel does for a field that's already first/last/alone),
  // then a divider, then the "Move to [other zone]" entries, then a final divider, then "Remove
  // Field" (capitalized to match — this used to read "Remove field" with no divider before it and
  // sat at the TOP of the menu instead, neither of which matched the real one), then "Field
  // Settings…" last (PLAN 6, 2026-10-13) — opens a sub-panel instead of closing the menu, switching
  // this pill's own `panel` state rather than calling onSelect+closeMenu like every other item.
  // Excluded for Σ Values (not a real field, nothing to configure) and whenever the caller doesn't
  // pass `fieldSettings` at all (Columns today — Subtotals is Rows-only, see this component's own
  // prop comment).
  const menuItems = [
    { label: "Move Up", onSelect: onMoveUp, disabled: isFirst },
    { label: "Move Down", onSelect: onMoveDown, disabled: isLast },
    { label: "Move to Beginning", onSelect: onMoveToBeginning, disabled: isFirst },
    { label: "Move to End", onSelect: onMoveToEnd, disabled: isLast },
    ...(moveTargets.length > 0 ? [{ divider: true }] : []),
    ...moveTargets.map((target) => ({
      label: target.label,
      onSelect: () => onMoveTo(fullPath, target.zoneKey),
    })),
    { divider: true },
    { label: "Remove Field", onSelect: () => onRemove(fullPath) },
    ...((fieldSettings || fieldLayout || fieldShowNoData) && !isMeasuresPlaceholder
      ? [{ label: "Field Settings…", onSelect: () => setPanel("subtotals"), keepOpen: true }]
      : []),
  ];

  // overrides?. (not just fieldSettings?./fieldLayout?.) defends against a request/fieldLayouts
  // state restored from a localStorage snapshot saved before these fields existed — same "don't
  // trust an older persisted shape" caution as readPersistedBuilderState's own callers elsewhere;
  // caught live right after fixing the sibling bug below (this component's outer function never
  // actually forwarded fieldSettings to this child at all until now, so this code path had never
  // run against real browser state before).
  const currentSubtotal = fieldSettings?.overrides?.find((f) => f.field === fullPath);
  const currentOverride = currentSubtotal?.position ?? null;
  const currentFunctions = currentSubtotal?.functions ?? null;
  const effectivePosition = effectiveSubtotalPosition(currentOverride, fieldSettings?.reportDefault ?? "Off");
  const currentLayout = fieldLayout?.overrides?.find((f) => f.field === fullPath) ?? {};
  // Excel's Custom subtotal functions — the same 11 as "Summarize value field by" (2026-10-15).
  const SUBTOTAL_FUNCTIONS = SUMMARY_FUNCTIONS;

  return (
    <li
      draggable={Boolean(dragSourceKey)}
      onDragStart={(e) => onDragStart(e, fullPath)}
      onDragOver={dragSourceKey ? (e) => e.preventDefault() : undefined}
      onDrop={dragSourceKey ? (e) => onPillDrop(e, fullPath) : undefined}
      // min-w-0: a flex child (this <li>, inside the parent <ul>'s flex row) defaults to
      // min-width:auto, refusing to shrink below its own content's width even with w-full set —
      // the same trap as PivotBuilder.jsx's own flex-1/min-w-0 WHY comment, just on a flex item
      // instead of a grid track. Left unset, a long field name could force this pill (and the
      // whole row) wider than the card, spilling out past its right edge.
      className={`w-full min-w-0 ${dragSourceKey ? "cursor-grab active:cursor-grabbing" : ""}`}
    >
      {/* Excel's own field-button look: a full-width plain bordered button with a dropdown
          chevron, not a colored chip (flagged live, 2026-10-04, against a screenshot of real
          Excel's field list). A LEFT click on this arrow is the only way to open this menu — real
          Excel has no right-click path for it either (confirmed live, 2026-10-05), so there's no
          second, separately-positioned context menu here anymore. */}
      <Popover
        align="right"
        fullWidth
        // This panel is a short menu list, not a form — the shared default (min-w-[22rem], sized
        // for Popover's OTHER callers' real forms/selects) would make it far wider than this
        // narrow docked pane, pushing it past the card's own edge (caught live, 2026-10-05).
        panelClassName="!min-w-[11rem] !max-w-[14rem] !p-1"
        trigger={(toggle, open) => (
          <button
            type="button"
            onClick={() => {
              setPanel("menu");
              toggle();
            }}
            className={`flex w-full items-center justify-between gap-1 rounded border px-2 py-1.5 text-left text-xs font-medium text-ink shadow-sm hover:bg-slate-50 ${open ? "border-primary bg-primary-50 ring-1 ring-primary/40" : "border-slate-300 bg-white"}`}
          >
            {/* A field's Custom Name replaces its name in the Rows area too, like Excel's field list. */}
            {/* A field reads by its own caption, like Excel's areas ("Country Name"); full name on hover. */}
            <span className="min-w-0 truncate" title={formatDimensionName(fullPath)}>
              {currentLayout.customName || fieldCaption(fullPath)}
            </span>
            <ChevronDown className="h-3.5 w-3.5 shrink-0 text-muted" />
          </button>
        )}
      >
        {(closeMenu) =>
          panel === "subtotals" ? (
            // Real Excel's own "Field Settings…" is a two-tab modal (Subtotals & Filters / Layout &
            // Print) — rendered here as one scrolling panel instead of true tabs, matching this
            // app's existing single-scroll-panel convention for every other "Field Settings"-style
            // popover (ValueFieldSettingsPanel is the same shape, not tabbed). The Layout section
            // below deliberately does NOT reproduce Excel's literal "Outline form / Tabular form"
            // radio pair — this app's renderer treats Outline and Tabular identically (see
            // HtmlSnapshotRenderer.cs's own WHY comment), so a field's only real per-field choice is
            // whether it merges into a shared "compact" column with the next field or stays
            // standalone; a fake radio reproducing a distinction this app can't actually render
            // would be worse than an honestly narrower one.
            <div className="flex max-h-80 flex-col gap-1 overflow-y-auto text-sm">
              <button
                type="button"
                onClick={() => setPanel("menu")}
                className="mb-1 inline-flex items-center gap-1 px-2 text-xs font-medium text-muted hover:text-ink"
              >
                <ChevronDown className="h-3.5 w-3.5 rotate-90" /> Back
              </button>

              {fieldLayout && (
                <div className="px-2 pb-2">
                  <label className="mb-1 block text-xs font-bold uppercase tracking-wide text-muted">Custom Name</label>
                  <input
                    type="text"
                    className="w-full rounded border border-slate-300 px-2 py-1 text-sm"
                    placeholder={fieldCaption(fullPath)}
                    value={currentLayout.customName ?? ""}
                    onChange={(e) => fieldLayout.onChange(fullPath, { customName: e.target.value || null })}
                  />
                </div>
              )}

              {fieldSettings && (
                <div className="px-2 pb-2">
                  <p className="pb-1 text-xs font-bold uppercase tracking-wide text-muted">Subtotals</p>
                  {[
                    { key: "automatic", label: "Automatic", active: effectivePosition !== "Off" && !currentFunctions },
                    { key: "none", label: "None", active: effectivePosition === "Off" },
                    { key: "custom", label: "Custom", active: effectivePosition !== "Off" && Boolean(currentFunctions) },
                  ].filter((opt) => !(fieldSettings.noCustom && opt.key === "custom")).map((opt) => (
                    <label key={opt.key} className="flex cursor-pointer items-center gap-2 rounded-lg py-1.5 text-left hover:bg-primary/10">
                      <input
                        type="radio"
                        className="h-3.5 w-3.5 accent-primary"
                        checked={opt.active}
                        onChange={() => {
                          // Real Excel: Automatic/Custom always turn THIS field's subtotal on, None
                          // always turns it off — whatever the report-wide Design > Subtotals choice
                          // was. So an override is only stored when the field differs from the report
                          // default; when the default is Off, turning a field on needs an explicit
                          // position (Bottom, same as Custom has always used here).
                          const reportOff = fieldSettings.reportDefault === "Off";
                          const onPosition = reportOff ? (currentOverride && currentOverride !== "Off" ? currentOverride : "Bottom") : currentOverride === "Off" ? null : currentOverride;
                          if (opt.key === "automatic") {
                            fieldSettings.onChange(fullPath, { position: onPosition, functions: null });
                          } else if (opt.key === "none") {
                            fieldSettings.onChange(fullPath, { position: reportOff ? null : "Off", functions: null });
                          } else {
                            fieldSettings.onChange(fullPath, { position: onPosition, functions: currentFunctions ?? ["Sum"] });
                          }
                        }}
                      />
                      {opt.label}
                    </label>
                  ))}
                  {currentFunctions && effectivePosition !== "Off" && (
                    <div className="ml-5 flex flex-col gap-0.5 border-l border-slate-100 pl-2">
                      {SUBTOTAL_FUNCTIONS.map((fn) => (
                        <label key={fn} className="flex cursor-pointer items-center gap-2 py-0.5 text-left text-xs">
                          <input
                            type="checkbox"
                            className="h-3.5 w-3.5 accent-primary"
                            checked={currentFunctions.includes(fn)}
                            onChange={(e) => {
                              const next = e.target.checked ? [...currentFunctions, fn] : currentFunctions.filter((f) => f !== fn);
                              fieldSettings.onChange(fullPath, { functions: next });
                            }}
                          />
                          {aggregationLabel(fn)}
                        </label>
                      ))}
                    </div>
                  )}
                </div>
              )}

              {/* nameOnly (Columns): just the Custom Name — these layout options are Rows-only. */}
              {fieldLayout && !fieldLayout.nameOnly && (
                <div className="px-2 pb-2">
                  <p className="pb-1 text-xs font-bold uppercase tracking-wide text-muted">Layout</p>
                  {(() => {
                    // Excel shows a field's EFFECTIVE state — what it inherits from the report-wide
                    // layout when it has no override of its own (2026-10-14: an inherited Compact
                    // used to show unchecked). Setting a value equal to the inherited one stores no
                    // override, so the field keeps following the report-wide setting afterward.
                    const inheritedCompact = fieldLayout.reportLayout === "Compact";
                    const effectiveCompact = currentLayout.compact ?? inheritedCompact;
                    // A compact-merged field ignores the report-wide Repeat All Item Labels (same as
                    // the renderers — Excel greys it out in Compact Form).
                    const inheritedRepeat = effectiveCompact ? false : Boolean(fieldLayout.reportRepeatItemLabels);
                    return [
                      { key: "compact", label: "Compact form (merge with next field)", value: effectiveCompact, inherited: inheritedCompact },
                      { key: "repeatItemLabels", label: "Repeat item labels", value: currentLayout.repeatItemLabels ?? inheritedRepeat, inherited: inheritedRepeat },
                      { key: "insertBlankLine", label: "Insert blank line after each item", value: currentLayout.insertBlankLine ?? false, inherited: false },
                    ];
                  })().map((opt) => (
                    <label key={opt.key} className="flex cursor-pointer items-center gap-2 py-1 text-left">
                      <input
                        type="checkbox"
                        className="h-3.5 w-3.5 accent-primary"
                        checked={opt.value}
                        onChange={(e) => fieldLayout.onChange(fullPath, { [opt.key]: e.target.checked === opt.inherited ? null : e.target.checked })}
                      />
                      {opt.label}
                    </label>
                  ))}
                </div>
              )}

              {fieldShowNoData && (
                <div className="px-2 pb-2">
                  <label className="flex cursor-pointer items-center gap-2 py-1 text-left">
                    <input
                      type="checkbox"
                      className="h-3.5 w-3.5 accent-primary"
                      checked={Boolean(fieldShowNoData.fields?.includes(fullPath))}
                      onChange={(e) => fieldShowNoData.onChange(fullPath, e.target.checked)}
                    />
                    Show items with no data
                  </label>
                  {/* EXPERIMENTAL (PLAN 7, 2026-10-13) — this query-affecting checkbox's generated
                      MDX hasn't been run against the real cube yet; see the plan file's own WHY
                      comment. Flagged inline, not just in a WHY comment, so it's visible to whoever
                      is testing it live, not just someone reading the source. */}
                  <p className="pt-1 text-[11px] leading-snug text-muted">Experimental — verify against the real cube before relying on this.</p>
                </div>
              )}
            </div>
          ) : (
            <div className="flex flex-col gap-1 text-sm">
              {menuItems.map((item, i) =>
                item.divider ? (
                  <div key={`divider-${i}`} className="my-1 border-t border-slate-100" />
                ) : (
                  <button
                    key={item.label}
                    type="button"
                    disabled={item.disabled}
                    onClick={() => {
                      item.onSelect();
                      if (!item.keepOpen) closeMenu();
                    }}
                    className="rounded-lg px-2 py-1.5 text-left hover:bg-primary/10 disabled:cursor-not-allowed disabled:text-muted/50 disabled:hover:bg-transparent"
                  >
                    {item.label}
                  </button>
                ),
              )}
            </div>
          )
        }
      </Popover>
    </li>
  );
}

const selectStyles = {
  menuPortal: (base) => ({ ...base, zIndex: 10000 }),
  control: (base) => ({
    ...base,
    borderRadius: "8px", // --ds-radius-control
    borderColor: "#e2e8f0",
    "&:hover": { borderColor: "#126e5c" },
  }),
  multiValue: (base) => ({
    ...base,
    backgroundColor: "#126e5c",
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
    "&:hover": { backgroundColor: "#0d594b" },
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
  icon: Icon = null,
  placeholderAdd = "Choose dimension and field(s), then add — all matching paths are added at once.",
  addButtonLabel = "Add",
  dragSourceKey = null,
  onFieldDropped = null,
  // true when the caller wraps several of these in one shared card (Columns + Rows inside the
  // Fields strip's merged "real drop zones" card, 2026-09-27) — drops this component's own
  // border/shadow/background/padding so it doesn't nest a card inside a card, keeping only the
  // flex layout its pill list relies on.
  bare = false,
  // "right" for whichever instance ends up rightmost on screen (Rows, once Filters/Columns/Rows
  // merged into one card next to the Values card pinned at the strip's right edge) — Popover's
  // default "left" expands rightward with nowhere to go there and gets clipped by the browser edge
  // (same overflow bug already fixed once for the Values pill's own popover, caught live again here
  // for this one, 2026-09-27).
  align = "left",
  // Rejects a cross-zone drop whose field isn't one this zone can actually hold — without this, a
  // measure dragged from the Values pill landed here unguarded (Values is the only zone that had
  // its own isFieldAccepted check) and corrupted state: `onChange([...value, measureName])` just
  // pushed a measure name into the Rows/Columns array, a dimension-only list (caught live,
  // 2026-09-28). Defaults to accepting everything so existing callers that never needed this are
  // unaffected.
  isFieldAccepted = () => true,
  // Extra "Move to [zone]" entries for each pill's own dropdown menu, e.g.
  // [{ zoneKey: "filters", label: "Move to Report Filter" }] — including "values" is fine, a
  // dimension field CAN move there (it just becomes a Count/Sum measure, same as dragging it in —
  // see FieldPicker.jsx's own moveField). Defaults to none so callers that don't pass it just get
  // the Move Up/Down/Beginning/End + Remove Field items below.
  contextMenuTargets = [],
  // Real Excel's per-field "Field Settings…" (2026-10-13) — forwarded straight through to each
  // PlacedFieldPill below. NOTE: this outer component never actually destructured/forwarded
  // fieldSettings before this fix (PLAN 6 Phase 2, 2026-10-13) — it was passed in from FieldPicker.jsx
  // but silently dropped here, so the Phase 1 "Field Settings…" menu item never actually appeared;
  // caught while adding fieldLayout alongside it.
  fieldSettings,
  fieldLayout,
  fieldShowNoData,
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

  // Backs the pill menu's "Move Up"/"Move Down"/"Move to Beginning"/"Move to End" items (confirmed
  // live 2026-10-13 against a real Excel screenshot of this exact menu — this app only had
  // drag-to-reorder before, not these explicit commands). newIndex is clamped so "Move Up" from the
  // very top or "Move Down" from the very bottom are no-ops rather than throwing — PlacedFieldPill
  // disables those buttons anyway, this is just defense in depth.
  const moveWithinList = (fullPath, newIndex) => {
    const without = value.filter((v) => v !== fullPath);
    const clamped = Math.max(0, Math.min(newIndex, without.length));
    onChange([...without.slice(0, clamped), fullPath, ...without.slice(clamped)]);
  };

  const [dragOver, setDragOver] = useState(false);
  const dragActive = useFieldDragActive();

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
    // Stops the drop from bubbling up to PivotBuilder's own page-level "dropped outside every
    // zone = remove the field" fallback (Excel's own behavior: drag a field out of the Field List
    // entirely, anywhere on the page, and it's removed) — this IS a recognized zone, so the drop
    // ends here whether or not it's actually accepted below, same as Excel doing nothing (not
    // removing) when you drop ON a zone but the field doesn't belong there.
    e.stopPropagation();
    setDragOver(false);
    const payload = readFieldDragPayload(e);
    if (payload && payload.source !== dragSourceKey && isFieldAccepted(payload.field)) {
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
    } else if (onFieldDropped && isFieldAccepted(payload.field)) {
      onFieldDropped(payload.field, payload.source, dragSourceKey);
    }
  };

  return (
    // flex h-full flex-col: this box sits in a flex row (FieldPicker's "Fields" strip) that
    // stretches every card to match the tallest sibling — without this, the <ul> drop target below
    // stayed sized to its own pill content, leaving genuine leftover white space inside a stretched
    // card that LOOKED droppable but wasn't (caught live, 2026-09-23). flex-1 on the <ul> below is
    // what actually claims that leftover space.
    <div className={`flex h-full flex-col p-2 ${bare ? "" : "rounded-xl border border-slate-200 bg-white shadow-sm"}`}>
      {/* min-w-0 + truncate: in a narrow docked pane (1024px window, 2026-10-25) the label shrinks
          instead of pushing the "+" button out past the card's edge. */}
      <div className="mb-1.5 flex min-w-0 items-center justify-between gap-1">
        <span className="flex min-w-0 items-center gap-1.5 truncate text-xs font-bold uppercase tracking-wide text-muted">
          {Icon && <Icon className="h-3.5 w-3.5 shrink-0" />}
          {label} {value.length > 0 && <span className="text-primary">({value.length})</span>}
        </span>
        <Popover
          align={align}
          trigger={(toggle, open) => (
            <button
              type="button"
              onClick={toggle}
              className={`flex h-6 w-6 shrink-0 items-center justify-center rounded-full ${open ? "bg-primary text-white" : "bg-primary/15 text-primary hover:bg-primary/25"}`}
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
          of the droppable area (the <ul> itself still gets the drop handlers).
          min-h-[6.5rem] — kept matched to FieldPicker.jsx's own ZoneDropArea (Filters/Values), which
          this component's Columns/Rows boxes must stay the same height as ("الاربع مربعات مش نفس
          الطول"). Both were cut to 2.75rem chasing a whole-page scrollbar bug that, in the end,
          turned out to be caused by AppShell.jsx's own padding instead — once that was fixed
          directly there, both got restored past the original 6rem, to 6.5rem, on request. */}
      {/* Same drag feedback as the Filters/Values areas (2026-10-21): every area outlined while a field
          is in flight, the one under the pointer soft gold with "Drop to add", and no flicker when
          the pointer crosses a pill inside it. */}
      <ul
        data-zone={dragSourceKey}
        className={`flex min-h-[6.5rem] flex-1 flex-wrap content-start gap-1.5 rounded-md transition-colors ${value.length === 0 ? "items-center" : "items-start"} ${
          dragOver ? "bg-primary-soft/70 ring-2 ring-inset ring-primary" : dragActive ? "bg-slate-50 outline-dashed outline-1 -outline-offset-1 outline-slate-300" : ""
        } ${value.length === 0 ? "border border-dashed border-slate-300 px-2" : ""}`}
        onDragOver={dragSourceKey && onFieldDropped ? handleDragOver : undefined}
        onDragLeave={dragSourceKey && onFieldDropped ? (e) => !e.currentTarget.contains(e.relatedTarget) && setDragOver(false) : undefined}
        onDrop={dragSourceKey && onFieldDropped ? handleDrop : undefined}
      >
        {value.length === 0 && <span className={`text-xs ${dragOver ? "font-medium text-primary-strong" : "text-slate-400"}`}>{dragOver ? "Drop to add" : "Drop a field here"}</span>}
        {value.map((fullPath, index) => (
          <PlacedFieldPill
            key={fullPath}
            fullPath={fullPath}
            dragSourceKey={dragSourceKey}
            onDragStart={handleDragStart}
            onPillDrop={handlePillDrop}
            onRemove={remove}
            contextMenuTargets={contextMenuTargets}
            onMoveTo={(field, toZone) => dragSourceKey && onFieldDropped && onFieldDropped(field, dragSourceKey, toZone)}
            isFirst={index === 0}
            isLast={index === value.length - 1}
            onMoveUp={() => moveWithinList(fullPath, index - 1)}
            onMoveDown={() => moveWithinList(fullPath, index + 1)}
            onMoveToBeginning={() => moveWithinList(fullPath, 0)}
            onMoveToEnd={() => moveWithinList(fullPath, value.length - 1)}
            fieldSettings={fieldSettings}
            fieldLayout={fieldLayout}
            fieldShowNoData={fieldShowNoData}
          />
        ))}
      </ul>
    </div>
  );
}
