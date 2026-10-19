import { useState } from "react";
import { Sigma, Rows3, Rows4, Calculator, ListPlus, LayoutGrid, CalendarRange, Settings2, SquarePlus, Heading, PanelRight, TableProperties, Eraser, Files } from "lucide-react";
import Popover from "../forms/Popover";
import { CalculatedFieldDialog, CalculatedItemDialog } from "./FieldPicker";
import { MEASURES_PLACEHOLDER } from "../../utils/measuresPlacement";
import { applyReportSubtotals } from "../../utils/subtotals";
import { isCompleteCalculatedField } from "../../utils/calculatedFields";

// Excel's "Insert Slicers" / "Insert Timelines" dialogs (2026-10-14, against screenshots): a checkbox
// list, so SEVERAL fields can be picked at once; OK stays disabled until one is ticked; Cancel. The
// cube has thousands of fields, so a search box sits above the list.
function InsertFieldsDialog({ title, fields, emptyText, onOk, onCancel }) {
  const [checked, setChecked] = useState([]);
  const [search, setSearch] = useState("");
  const visible = fields.filter((d) => d.displayName.toLowerCase().includes(search.toLowerCase()));
  const toggle = (field) => setChecked((current) => (current.includes(field) ? current.filter((f) => f !== field) : [...current, field]));
  return (
    <div className="w-60 space-y-2">
      <p className="text-sm font-semibold text-ink">{title}</p>
      <input type="text" autoFocus value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search fields…" className="input-field w-full py-1.5 text-sm" />
      <div className="h-56 overflow-y-auto rounded-md border border-slate-200 bg-white p-1">
        {visible.length === 0 ? (
          <p className="px-1 py-2 text-xs text-muted">{fields.length === 0 ? emptyText : "No matching fields."}</p>
        ) : (
          visible.map((d) => (
            <label key={d.field} className="flex cursor-pointer items-center gap-2 rounded px-1 py-1 text-sm hover:bg-gold/10">
              <input type="checkbox" className="h-3.5 w-3.5 accent-gold" checked={checked.includes(d.field)} onChange={() => toggle(d.field)} />
              <span className="truncate">{d.displayName}</span>
            </label>
          ))
        )}
      </div>
      <div className="flex justify-end gap-2">
        <button type="button" disabled={checked.length === 0} onClick={() => onOk(checked)} className="btn-primary px-3 py-1 text-xs disabled:cursor-not-allowed disabled:opacity-50">
          OK
        </button>
        <button type="button" onClick={onCancel} className="btn-secondary px-3 py-1 text-xs">
          Cancel
        </button>
      </div>
    </div>
  );
}

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
//
// panelClassName widens the panel for a bigger dialog (Popover's default caps it at 26rem, and a
// dialog wider than that spilled its buttons out of the card — caught live 2026-10-14). Every ribbon
// panel also opens a little lower (!mt-6), clear of the group caption under the button ("Calculations").
function RibbonButton({ icon: Icon, label, badge, panelClassName = "", children }) {
  return (
    <Popover
      panelClassName={`!mt-6 ${panelClassName}`}
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

// pivotOptions/onPivotOptionsChange/onOpenPivotTableOptions: the PivotTable group (name + Options);
// omitted, that group isn't drawn.
// fieldListShown/onToggleFieldList: Excel's "Field List" toggle in the Show group; omitted, no button.
// layout: { reportLayout, repeatItemLabels, fieldLayouts, onReportLayoutChange, onRepeatItemLabelsChange,
// onFieldLayoutsChange } — the Design tab's Report Layout / Blank Rows (2026-10-20); omitted, not drawn.
export default function PivotToolbar({ request, onChange, measureOptions, dimensions = [], pivotOptions = null, onPivotOptionsChange = null, onOpenPivotTableOptions = null, fieldListShown = true, onToggleFieldList = null, layout = null, actions = null }) {
  const [tab, setTab] = useState("analyze");
  // The name box's text while it's being typed (null = show the saved name); applied on blur/Enter.
  const [nameDraft, setNameDraft] = useState(null);

  // Insert Slicer: adds the field to slicerFields (the Slicer widget renders above the table from
  // there — see PivotGrid.jsx) and, if nothing already filters this field, gives it a fresh "(All)"
  // Members-mode entry in request.filters — same default an empty Report Filter row starts from, so
  // the Slicer opens with every member unselected rather than accidentally inheriting some other
  // mode's leftover state.
  //
  // Several at once since 2026-10-14 (Excel's checkbox dialog) — one onChange for all of them.
  const withFilterEntries = (filters, fields) =>
    fields.reduce((acc, field) => (acc.some((f) => f.field === field) ? acc : [...acc, { field, mode: "Members", includedMembers: [] }]), filters);
  const addSlicers = (fields) => {
    const added = fields.filter((f) => !request.slicerFields.includes(f));
    onChange({ ...request, slicerFields: [...request.slicerFields, ...added], filters: withFilterEntries(request.filters, added) });
  };

  // Insert Timeline — same relationship to request.filters as addSlicers above, for the date-range
  // widget. Like Excel's, the list holds DATE fields only — the cube metadata's isDate (its level's
  // DBTYPE_DATE, see the backend's CubeField.IsDate).
  const addTimelines = (fields) => {
    const added = fields.filter((f) => !request.timelineFields.includes(f));
    onChange({ ...request, timelineFields: [...request.timelineFields, ...added], filters: withFilterEntries(request.filters, added) });
  };

  const availableForSlicer = dimensions.filter((d) => !request.slicerFields.includes(d.field));
  const availableForTimeline = dimensions.filter((d) => d.isDate && !request.timelineFields.includes(d.field));

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
      {/* overflow-x-auto (2026-10-19): on a narrow window the groups scroll sideways instead of
          being cut off (the Calculations/Show groups used to vanish). Dropdown panels render in a
          portal, so this doesn't clip them. */}
      <div className="custom-scrollbar flex items-stretch gap-0 overflow-x-auto bg-slate-50 px-2 py-1">
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
                // A drilldown hierarchy has subtotals of its own levels (Drilldown v2, 2026-10-14).
                // Since 2026-10-15 two Columns fields count too (their "2024 Total" columns).
                const needsMoreRowFields =
                  request.rows.filter((f) => f !== MEASURES_PLACEHOLDER).length < 2 &&
                  !request.rowsHierarchyDrilldown &&
                  request.columns.filter((f) => f !== MEASURES_PLACEHOLDER).length < 2;
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
                    {needsMoreRowFields && <p className="px-2 pt-1 text-xs text-muted">Needs 2+ Row or Column fields.</p>}
                  </div>
                );
              }}
            </RibbonButton>
            {layout && (
              <>
                {/* Excel's Design › Report Layout dropdown (2026-10-20): the three forms plus Repeat / Do
                    Not Repeat Item Labels — the same settings as the layout box under the field list. */}
                <RibbonButton icon={TableProperties} label="Report Layout" badge={0}>
                  {(close) => (
                    <div className="flex flex-col gap-1 text-sm">
                      {[
                        { value: "Compact", label: "Show in Compact Form" },
                        { value: "Outline", label: "Show in Outline Form" },
                        { value: "Tabular", label: "Show in Tabular Form" },
                      ].map((opt) => (
                        <button
                          key={opt.value}
                          type="button"
                          onClick={() => {
                            layout.onReportLayoutChange(opt.value);
                            close();
                          }}
                          className={`rounded-lg px-2 py-1.5 text-left ${layout.reportLayout === opt.value ? "bg-gold/20 font-semibold text-gold-strong" : "hover:bg-gold/10"}`}
                        >
                          {opt.label}
                        </button>
                      ))}
                      <div className="my-1 border-t border-slate-100" />
                      {[
                        { value: true, label: "Repeat All Item Labels" },
                        { value: false, label: "Do Not Repeat Item Labels" },
                      ].map((opt) => (
                        <button
                          key={opt.label}
                          type="button"
                          // Excel greys these out in Compact Form (one column — nothing to repeat into).
                          disabled={layout.reportLayout === "Compact"}
                          onClick={() => {
                            layout.onRepeatItemLabelsChange(opt.value);
                            close();
                          }}
                          className={`rounded-lg px-2 py-1.5 text-left disabled:cursor-not-allowed disabled:opacity-40 ${
                            layout.reportLayout !== "Compact" && layout.repeatItemLabels === opt.value ? "bg-gold/20 font-semibold text-gold-strong" : "hover:bg-gold/10"
                          }`}
                        >
                          {opt.label}
                        </button>
                      ))}
                    </div>
                  )}
                </RibbonButton>
                {/* Excel's Design › Blank Rows: a blank line after each item of every row field but the
                    innermost (that's where Excel puts them) — the per-field "Insert blank line after
                    each item" setting, set or cleared for all of them at once. */}
                <RibbonButton icon={Rows4} label="Blank Rows" badge={layout.fieldLayouts.some((f) => f.insertBlankLine) ? 1 : 0}>
                  {(close) => {
                    const outerFields = request.rows.filter((f) => f !== MEASURES_PLACEHOLDER).slice(0, -1);
                    const setBlankLines = (on) => {
                      const byField = new Map(layout.fieldLayouts.map((f) => [f.field, f]));
                      const touched = on ? outerFields : layout.fieldLayouts.map((f) => f.field);
                      touched.forEach((field) => byField.set(field, { ...(byField.get(field) ?? { field }), insertBlankLine: on }));
                      // A field left with nothing but defaults is dropped, like the field settings do.
                      const next = [...byField.values()].filter((f) => f.customName || f.compact != null || f.repeatItemLabels != null || f.insertBlankLine);
                      layout.onFieldLayoutsChange(next);
                      close();
                    };
                    return (
                      <div className="flex flex-col gap-1 text-sm">
                        <button
                          type="button"
                          disabled={outerFields.length === 0}
                          onClick={() => setBlankLines(true)}
                          className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10 disabled:cursor-not-allowed disabled:opacity-40"
                        >
                          Insert Blank Line after Each Item
                        </button>
                        <button type="button" onClick={() => setBlankLines(false)} className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10">
                          Remove Blank Line after Each Item
                        </button>
                        {outerFields.length === 0 && <p className="px-2 pt-1 text-xs text-muted">Needs 2+ Row fields.</p>}
                      </div>
                    );
                  }}
                </RibbonButton>
              </>
            )}
          </RibbonGroup>
        )}

        {/* Excel's first Analyze group: the PivotTable's name (editable in place) and the Options
            button that opens the PivotTable Options dialog (2026-10-18). */}
        {tab === "analyze" && pivotOptions && (
          <RibbonGroup label="PivotTable">
            <div className="flex h-16 w-40 flex-col justify-center gap-1 px-1">
              <label className="text-[11px] text-muted" htmlFor="pivot-table-name">
                PivotTable Name:
              </label>
              <input
                id="pivot-table-name"
                type="text"
                value={nameDraft ?? pivotOptions.name}
                onChange={(e) => setNameDraft(e.target.value)}
                onBlur={() => {
                  if (nameDraft !== null) onPivotOptionsChange({ ...pivotOptions, name: nameDraft.trim() || pivotOptions.name });
                  setNameDraft(null);
                }}
                onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()}
                className="rounded border border-slate-300 bg-white px-1.5 py-0.5 text-xs focus:outline-none focus:ring-1 focus:ring-gold"
              />
            </div>
            <button
              type="button"
              onClick={onOpenPivotTableOptions}
              className="relative flex h-16 w-16 shrink-0 flex-col items-center justify-center gap-1 rounded px-1 text-ink hover:bg-black/5"
            >
              <Settings2 className="h-5 w-5 text-ink/70" />
              <span className="text-center text-[11px] leading-tight">Options</span>
            </button>
            {/* Excel's Options ▸ Show Report Filter Pages (2026-10-20): one copy of the pivot per item
                of a Report Filter field — here, one sheet each in a downloaded workbook. */}
            {actions && (
              <RibbonButton icon={Files} label="Report Filter Pages" badge={0}>
                {(close) => (
                  <div className="w-60 space-y-2 text-sm">
                    <p className="font-semibold text-ink">Show Report Filter Pages</p>
                    {actions.reportFilterFields.length === 0 ? (
                      <p className="text-xs text-muted">Add a field to the Filters area (and run the pivot) first.</p>
                    ) : (
                      <>
                        <p className="text-xs text-muted">Show all report filter pages of:</p>
                        <div className="flex flex-col gap-1">
                          {actions.reportFilterFields.map((f) => (
                            <button
                              key={f.field}
                              type="button"
                              onClick={() => {
                                actions.onShowReportFilterPages(f.field);
                                close();
                              }}
                              className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10"
                            >
                              {f.label}
                            </button>
                          ))}
                        </div>
                      </>
                    )}
                    {actions.filterPagesStatus && <p className="text-xs text-muted">{actions.filterPagesStatus}</p>}
                  </div>
                )}
              </RibbonButton>
            )}
          </RibbonGroup>
        )}

        {/* Excel's Analyze › Actions › Clear (2026-10-20). (Select and Move PivotTable act on worksheet
            cells — nothing to map them to on a web page.) */}
        {tab === "analyze" && actions && (
          <RibbonGroup label="Actions">
            <RibbonButton icon={Eraser} label="Clear" badge={0}>
              {(close) => (
                <div className="flex flex-col gap-1 text-sm">
                  <button
                    type="button"
                    onClick={() => {
                      actions.onClearAll();
                      close();
                    }}
                    className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10"
                  >
                    Clear All
                  </button>
                  <button
                    type="button"
                    disabled={request.filters.length === 0}
                    onClick={() => {
                      actions.onClearFilters();
                      close();
                    }}
                    className="rounded-lg px-2 py-1.5 text-left hover:bg-gold/10 disabled:cursor-not-allowed disabled:opacity-40"
                  >
                    Clear Filters
                  </button>
                </div>
              )}
            </RibbonButton>
          </RibbonGroup>
        )}

        {tab === "analyze" && (
          <RibbonGroup label="Filter">
            <RibbonButton icon={LayoutGrid} label="Insert Slicer" badge={0}>
              {(close) => (
                <InsertFieldsDialog
                  title="Insert Slicers"
                  fields={availableForSlicer}
                  emptyText="No fields available."
                  onOk={(fields) => {
                    addSlicers(fields);
                    close();
                  }}
                  onCancel={close}
                />
              )}
            </RibbonButton>
            <RibbonButton icon={CalendarRange} label="Insert Timeline" badge={0}>
              {(close) => (
                <InsertFieldsDialog
                  title="Insert Timelines"
                  fields={availableForTimeline}
                  emptyText="No date fields in the selected tables."
                  onOk={(fields) => {
                    addTimelines(fields);
                    close();
                  }}
                  onCancel={close}
                />
              )}
            </RibbonButton>
          </RibbonGroup>
        )}

        {tab === "analyze" && (
          <RibbonGroup label="Calculations">
            <RibbonButton icon={Calculator} label="Calculated Field" badge={request.calculatedFields.length} panelClassName="!max-w-[34rem]">
              {(close) => (
                <CalculatedFieldDialog
                  calculatedFields={request.calculatedFields}
                  measureOptions={measureOptions}
                  onApply={(calculatedFields, addedName) => {
                    // A Value referencing a deleted calculated field becomes invalid; a NEW one is
                    // added to Values, like Excel does.
                    const validFieldNames = new Set(measureOptions.map((m) => m.value).concat(calculatedFields.filter(isCompleteCalculatedField).map((f) => f.name)));
                    const kept = request.values.filter((v) => validFieldNames.has(v.field) || v.field.startsWith("["));
                    const values = addedName && !kept.some((v) => v.field === addedName) ? [...kept, { field: addedName, showValuesAs: "Normal" }] : kept;
                    onChange({ ...request, calculatedFields, values });
                  }}
                  onClose={close}
                />
              )}
            </RibbonButton>
            <RibbonButton icon={ListPlus} label="Calculated Item" badge={request.calculatedItems.length} panelClassName="!max-w-[34rem]">
              {(close) => (
                <CalculatedItemDialog
                  calculatedItems={request.calculatedItems}
                  rows={request.rows}
                  columns={request.columns}
                  onApply={(calculatedItems) => onChange({ ...request, calculatedItems })}
                  onClose={close}
                />
              )}
            </RibbonButton>
          </RibbonGroup>
        )}

        {/* Excel's "Show" group: the same two settings as PivotTable Options > Display, as
            pressed/unpressed toggles (2026-10-19). */}
        {tab === "analyze" && pivotOptions && (
          <RibbonGroup label="Show">
            {/* Excel's "Field List" toggle — hides the field pane so the grid gets the full width. */}
            {onToggleFieldList && (
              <button
                type="button"
                aria-pressed={fieldListShown}
                onClick={onToggleFieldList}
                className={`flex h-16 w-20 shrink-0 flex-col items-center justify-center gap-1 rounded px-1 text-ink ${fieldListShown ? "bg-gold/15 ring-1 ring-gold/40" : "hover:bg-black/5"}`}
              >
                <PanelRight className="h-5 w-5 text-ink/70" />
                <span className="text-center text-[11px] leading-tight">Field List</span>
              </button>
            )}
            {[
              { key: "showExpandButtons", label: "+/- Buttons", icon: SquarePlus },
              { key: "showFieldCaptions", label: "Field Headers", icon: Heading },
            ].map(({ key, label, icon: Icon }) => {
              const on = pivotOptions[key] !== false;
              return (
                <button
                  key={key}
                  type="button"
                  aria-pressed={on}
                  onClick={() => onPivotOptionsChange({ ...pivotOptions, [key]: !on })}
                  className={`flex h-16 w-20 shrink-0 flex-col items-center justify-center gap-1 rounded px-1 text-ink ${on ? "bg-gold/15 ring-1 ring-gold/40" : "hover:bg-black/5"}`}
                >
                  <Icon className="h-5 w-5 text-ink/70" />
                  <span className="text-center text-[11px] leading-tight">{label}</span>
                </button>
              );
            })}
          </RibbonGroup>
        )}
      </div>
    </div>
  );
}
