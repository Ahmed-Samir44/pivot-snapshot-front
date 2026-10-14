import { useState } from "react";
import { Sigma, Rows3, Calculator, ListPlus, LayoutGrid, CalendarRange } from "lucide-react";
import Popover from "../forms/Popover";
import { CalculatedFieldsEditor, CalculatedItemsEditor } from "./FieldPicker";
import { MEASURES_PLACEHOLDER } from "../../utils/measuresPlacement";
import { applyReportSubtotals } from "../../utils/subtotals";

// A real Excel RIBBON, not just a toolbar in the ribbon's general spot — a plain restyled button
// row was tried first and rejected live, 2026-10-04 ("مشكلتي انه مش زي الـ excel"), after an
// earlier pass had already moved these controls out of FieldPicker's side pane into this spot
// (confirmed as the right spot) but kept them as rounded pill buttons (confirmed as NOT right).
// Excel keeps Grand Totals/Subtotals on the Design tab and Calculated Field/Item on the
// PivotTable Analyze tab — two tabs, not one, since that's genuinely where Excel puts each.
// Calculated Field and Calculated Item are two separate top-level buttons in Excel's own
// "Calculations" ribbon group too (confirmed live, 2026-10-05, against a real Excel screenshot —
// an earlier version of this comment incorrectly assumed they were nested under one combined
// "Fields, Items & Sets" dropdown; they're genuinely two buttons side by side, matching exactly
// what's built here).
const TABS = [
  { key: "analyze", label: "PivotTable Analyze" },
  { key: "design", label: "Design" },
];

// Icon-over-label, bordered-group ribbon button — Excel's own button shape in every group, not a
// rounded pill. `badge` mirrors Excel's own subtle "this is currently set" affordance (Excel uses
// a pressed/highlighted button state for an active toggle; a small count badge reads the same way
// at a glance without needing real toggle-button semantics for a Popover trigger).
function RibbonButton({ icon: Icon, label, badge, children }) {
  return (
    <Popover
      trigger={(toggle) => (
        <button
          type="button"
          onClick={toggle}
          className="relative flex h-16 w-20 shrink-0 flex-col items-center justify-center gap-1 rounded px-1 text-ink hover:bg-black/5"
        >
          <Icon className="h-5 w-5 text-ink/70" />
          <span className="text-center text-[11px] leading-tight">{label}</span>
          {badge > 0 && (
            <span className="absolute right-1 top-1 flex h-4 w-4 items-center justify-center rounded-full bg-gold text-[10px] font-bold text-white">
              {badge}
            </span>
          )}
        </button>
      )}
    >
      {/* Every call site passes the popover's own content as plain JSX children
          (`<RibbonButton ...>{() => (...)}</RibbonButton>`), not a named `trigger=` attribute — this
          used to destructure a `trigger` prop that was therefore always undefined, so Popover's own
          `{children(close)}` call crashed with "TypeError: children is not a function" the instant
          ANY ribbon button was opened (every one of the 6 — Grand Totals/Subtotals too, not just the
          4 under the Analyze tab that got clicked first), taking the whole React tree down with it
          since nothing in this app has an error boundary (caught live, 2026-10-06, "بيودوني على صفحة
          فاضية" — a hard crash reads as a blank page with no on-screen error, and the browser's own
          console tab apparently didn't surface it either — this was found via the Vite dev server's
          own terminal-side HMR error log instead). */}
      {children}
    </Popover>
  );
}

// One ribbon "group" — a cluster of buttons with a shared bottom caption and a divider on its
// right, exactly Excel's own group convention (e.g. "Calculations", "Layout").
function RibbonGroup({ label, children }) {
  return (
    <div className="flex h-full flex-col items-stretch border-r border-slate-300/70 px-2 pt-1">
      <div className="flex flex-1 items-start gap-1">{children}</div>
      <p className="pb-1 text-center text-[10px] text-muted">{label}</p>
    </div>
  );
}

export default function PivotToolbar({ request, onChange, measureOptions, dimensions = [] }) {
  const [tab, setTab] = useState("analyze");
  const [slicerSearch, setSlicerSearch] = useState("");
  const [timelineSearch, setTimelineSearch] = useState("");

  // Insert Slicer: adds the field to slicerFields (the Slicer widget renders above the table from
  // there — see PivotGrid.jsx) and, if nothing already filters this field, gives it a fresh "(All)"
  // Members-mode entry in request.filters — same default an empty Report Filter row starts from, so
  // the Slicer opens with every member unselected rather than accidentally inheriting some other
  // mode's leftover state.
  const addSlicer = (field) => {
    if (request.slicerFields.includes(field)) return;
    const filters = request.filters.some((f) => f.field === field)
      ? request.filters
      : [...request.filters, { field, mode: "Members", includedMembers: [] }];
    onChange({ ...request, slicerFields: [...request.slicerFields, field], filters });
    setSlicerSearch("");
  };

  // Insert Timeline — same relationship to request.filters as addSlicer above, just for the
  // date-range widget instead of the button-grid one. No reliable "is this a date field" metadata
  // is exposed by the cube API this app already calls, so (same honest fallback Timeline.jsx itself
  // uses) ANY dimension can be picked here — Timeline.jsx simply shows "no date-shaped members" if
  // the chosen field's captions don't parse as dates, rather than silently failing or pretending to
  // restrict the list from a classification this app can't actually make reliably.
  const addTimeline = (field) => {
    if (request.timelineFields.includes(field)) return;
    const filters = request.filters.some((f) => f.field === field)
      ? request.filters
      : [...request.filters, { field, mode: "Members", includedMembers: [] }];
    onChange({ ...request, timelineFields: [...request.timelineFields, field], filters });
    setTimelineSearch("");
  };

  const availableForSlicer = dimensions.filter(
    (d) => !request.slicerFields.includes(d.field) && d.displayName.toLowerCase().includes(slicerSearch.toLowerCase()),
  );
  const availableForTimeline = dimensions.filter(
    (d) => !request.timelineFields.includes(d.field) && d.displayName.toLowerCase().includes(timelineSearch.toLowerCase()),
  );

  return (
    <div className="mb-3 overflow-hidden rounded-xl border border-slate-300/70 bg-white shadow-sm">
      {/* Tab strip — Excel's active tab merges into the ribbon surface below it (no border between
          them); inactive tabs sit on the title-bar-colored strip above that surface. */}
      <div className="flex gap-4 border-b border-slate-300/70 bg-slate-100 px-3 pt-1.5">
        {TABS.map((t) => (
          <button
            key={t.key}
            type="button"
            onClick={() => setTab(t.key)}
            className={`-mb-px border-b-2 px-1 pb-1.5 text-xs font-medium ${
              tab === t.key ? "border-gold bg-white text-gold-strong" : "border-transparent text-muted hover:text-ink"
            }`}
          >
            {t.label}
          </button>
        ))}
      </div>

      {/* Ribbon surface — Excel's light gray ribbon background, not a white card. */}
      <div className="flex items-stretch gap-0 bg-slate-50 px-2 py-1">
        {tab === "design" && (
          <RibbonGroup label="Layout">
            {/* Real Excel's own Design tab "Grand Totals"/"Subtotals" are each a 4-/3-way DROPDOWN,
                not a plain on/off checkbox (caught 2026-10-06, "عايز كل حاجة زي ال excel" — a
                checkbox was never real parity, just the closest quick approximation at the time).
                badge shows 1 whenever the current choice isn't the "nothing" option, same "this is
                currently set" affordance every other RibbonButton badge already uses. */}
            <RibbonButton icon={Sigma} label="Grand Totals" badge={request.grandTotals !== "Off" ? 1 : 0}>
              {/* `close` is Popover's own render-prop argument (see RibbonButton's own WHY comment
                  on `{children}` just being Popover's content function) — previously ignored
                  entirely, so picking an option never closed this dropdown; it just sat open until
                  an unrelated outside click happened to catch it, reading as "stuck" (caught live,
                  2026-10-13, "لما بدوس عليها بتهنج كدا شوية ومش بتتقفل" — real Excel's own dropdowns
                  close immediately on picking an option). */}
              {(close) => (
                <div className="flex flex-col gap-1 text-sm">
                  {[
                    { value: "Off", label: "Off" },
                    { value: "RowsAndColumns", label: "On for Rows and Columns" },
                    { value: "RowsOnly", label: "On for Rows Only" },
                    { value: "ColumnsOnly", label: "On for Columns Only" },
                  ].map((opt) => (
                    <button
                      key={opt.value}
                      type="button"
                      onClick={() => {
                        onChange({ ...request, grandTotals: opt.value });
                        close();
                      }}
                      className={`rounded-lg px-2 py-1.5 text-left ${request.grandTotals === opt.value ? "bg-gold/20 font-semibold text-gold-strong" : "hover:bg-gold/10"}`}
                    >
                      {opt.label}
                    </button>
                  ))}
                </div>
              )}
            </RibbonButton>
            <RibbonButton icon={Rows3} label="Subtotals" badge={request.subtotals !== "Off" ? 1 : 0}>
              {(close) => {
                // request.rows can contain the Σ Values sentinel as a plain element when it sits in
                // Rows — not a real grouping field, so it must not count toward the "2+ row fields"
                // gate (caught 2026-10-06, full codebase review: Rows = [Hospital, "[Measures]"]
                // enabled this with only ONE real field).
                const needsMoreRowFields = request.rows.filter((f) => f !== MEASURES_PLACEHOLDER).length < 2;
                return (
                  <div className="flex flex-col gap-1 text-sm">
                    {[
                      { value: "Off", label: "Do Not Show Subtotals" },
                      { value: "Bottom", label: "Show all Subtotals at Bottom of Group" },
                      { value: "Top", label: "Show all Subtotals at Top of Group" },
                    ].map((opt) => (
                      <button
                        key={opt.value}
                        type="button"
                        disabled={opt.value !== "Off" && needsMoreRowFields}
                        onClick={() => {
                          onChange(applyReportSubtotals(request, opt.value));
                          close();
                        }}
                        className={`rounded-lg px-2 py-1.5 text-left disabled:cursor-not-allowed disabled:opacity-40 ${
                          request.subtotals === opt.value ? "bg-gold/20 font-semibold text-gold-strong" : "hover:bg-gold/10"
                        }`}
                      >
                        {opt.label}
                      </button>
                    ))}
                    {needsMoreRowFields && <p className="px-2 pt-1 text-xs text-muted">Needs 2+ Row fields.</p>}
                  </div>
                );
              }}
            </RibbonButton>
          </RibbonGroup>
        )}

        {tab === "analyze" && (
          <RibbonGroup label="Filter">
            <RibbonButton icon={LayoutGrid} label="Insert Slicer" badge={0}>
              {/* Same ignored-close bug as Grand Totals/Subtotals above, fixed the same way —
                  picking a field here is a one-click action (it immediately adds the Slicer), so
                  the popover should close right after, matching real Excel. */}
              {(close) => (
                <div className="w-56">
                  <input
                    type="text"
                    autoFocus
                    value={slicerSearch}
                    onChange={(e) => setSlicerSearch(e.target.value)}
                    placeholder="Search fields…"
                    className="input-field mb-2 w-full py-1.5 text-sm"
                  />
                  <div className="max-h-56 overflow-y-auto">
                    {availableForSlicer.length === 0 ? (
                      <p className="px-1 py-2 text-xs text-muted">No matching fields.</p>
                    ) : (
                      availableForSlicer.map((d) => (
                        <button
                          key={d.field}
                          type="button"
                          onClick={() => {
                            addSlicer(d.field);
                            close();
                          }}
                          className="block w-full rounded-lg px-2 py-1.5 text-left text-sm hover:bg-gold/10"
                        >
                          {d.displayName}
                        </button>
                      ))
                    )}
                  </div>
                </div>
              )}
            </RibbonButton>
            <RibbonButton icon={CalendarRange} label="Insert Timeline" badge={0}>
              {/* Same ignored-close bug as the dropdowns above, same fix. */}
              {(close) => (
                <div className="w-56">
                  {/* Real Excel never lists every raw day for a date field either — it auto-groups
                      into Year/Quarter/Month/Day instead. Picking a Month/Year-level field here
                      (if the table has one) gets the same effect safely; picking a day-level field
                      directly works too, but fetches every distinct day up front — on a field with
                      years of daily data, that's the same unfiltered-high-cardinality cost as
                      anywhere else in this app (confirmed live, 2026-10-06, see DECISIONS.md). */}
                  <p className="mb-2 text-xs text-muted">Prefer a Month/Year-level field over a raw day field if the table has one.</p>
                  <input
                    type="text"
                    autoFocus
                    value={timelineSearch}
                    onChange={(e) => setTimelineSearch(e.target.value)}
                    placeholder="Search fields…"
                    className="input-field mb-2 w-full py-1.5 text-sm"
                  />
                  <div className="max-h-56 overflow-y-auto">
                    {availableForTimeline.length === 0 ? (
                      <p className="px-1 py-2 text-xs text-muted">No matching fields.</p>
                    ) : (
                      availableForTimeline.map((d) => (
                        <button
                          key={d.field}
                          type="button"
                          onClick={() => {
                            addTimeline(d.field);
                            close();
                          }}
                          className="block w-full rounded-lg px-2 py-1.5 text-left text-sm hover:bg-gold/10"
                        >
                          {d.displayName}
                        </button>
                      ))
                    )}
                  </div>
                </div>
              )}
            </RibbonButton>
          </RibbonGroup>
        )}

        {tab === "analyze" && (
          <RibbonGroup label="Calculations">
            <RibbonButton icon={Calculator} label="Calculated Field" badge={request.calculatedFields.length}>
              {() => (
                <CalculatedFieldsEditor
                  calculatedFields={request.calculatedFields}
                  measureOptions={measureOptions}
                  onChange={(calculatedFields) => {
                    // Same dangling-reference guard FieldPicker's Values picker already relies on:
                    // a Value referencing a calculated field becomes invalid the moment that field
                    // is removed or edited back to incomplete.
                    const validFieldNames = new Set(
                      measureOptions
                        .map((m) => m.value)
                        .concat(calculatedFields.filter((f) => f.name && f.leftField && f.rightField).map((f) => f.name)),
                    );
                    const values = request.values.filter((v) => validFieldNames.has(v.field));
                    onChange({ ...request, calculatedFields, values });
                  }}
                />
              )}
            </RibbonButton>
            <RibbonButton icon={ListPlus} label="Calculated Item" badge={request.calculatedItems.length}>
              {() => (
                <CalculatedItemsEditor
                  calculatedItems={request.calculatedItems}
                  rows={request.rows}
                  columns={request.columns}
                  onChange={(calculatedItems) => onChange({ ...request, calculatedItems })}
                />
              )}
            </RibbonButton>
          </RibbonGroup>
        )}
      </div>
    </div>
  );
}
