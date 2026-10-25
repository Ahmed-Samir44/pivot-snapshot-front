import { createContext, useContext, useEffect, useLayoutEffect, useRef, useState } from "react";
import { Redo2, Undo2, ChevronDown, ChevronUp, DecimalsArrowLeft, DecimalsArrowRight, DollarSign, Percent, Sigma, Rows3, Rows4, Calculator, ListPlus, LayoutGrid, CalendarRange, Settings2, SquarePlus, Heading, PanelRight, TableProperties, Eraser, Files } from "lucide-react";
import Popover from "../forms/Popover";
import { NUMBER_FORMAT_CHOICES, accountingStyle, changeDecimals, commaStyle, percentStyle, withFormatType } from "../../utils/numberFormatCommands";
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
    // w-full, not a fixed w-60: the ribbon popover is at least 22rem wide, so a 15rem dialog left an
    // empty strip down its right side (2026-10-25); filling it also shows more of long field names.
    <div className="w-full space-y-2">
      <p className="text-sm font-semibold text-ink">{title}</p>
      <input type="text" autoFocus value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search fields…" className="input-field w-full py-1.5 text-sm" />
      <div className="h-56 overflow-y-auto rounded-md border border-slate-200 bg-white p-1">
        {visible.length === 0 ? (
          <p className="px-1 py-2 text-xs text-muted">{fields.length === 0 ? emptyText : "No matching fields."}</p>
        ) : (
          visible.map((d) => (
            <label key={d.field} className="flex cursor-pointer items-center gap-2 rounded px-1 py-1 text-sm hover:bg-primary/10">
              <input type="checkbox" className="h-3.5 w-3.5 accent-primary" checked={checked.includes(d.field)} onChange={() => toggle(d.field)} />
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
const RIBBON_COLLAPSED_KEY = "pivotSnapshot.ribbonCollapsed";

const TABS = [
  { key: "home", label: "Home" },
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
// Excel's narrow ribbon (2026-10-20, from a screenshot at 120% zoom: the Show group had scrolled out
// of view): when the full ribbon doesn't fit, its large icon-over-label buttons become small
// icon-beside-label buttons stacked three high in each group — set by PivotToolbar, read here.
const RibbonCompact = createContext(false);
// The next step down (2026-10-25, a 1024px window still overflowed the small buttons): icon only,
// the command's name kept as its tooltip and for screen readers — like Excel's narrowest ribbon.
const RibbonIconOnly = createContext(false);

function ribbonButtonClass(compact, extra = "") {
  return compact
    ? `relative flex h-[21px] shrink-0 items-center gap-1.5 whitespace-nowrap rounded px-1.5 text-ink transition-colors ${extra}`
    : `relative flex h-16 w-20 shrink-0 flex-col items-center justify-center gap-1 rounded-md px-1 text-ink transition-colors ${extra}`;
}

// Each command's icon colour, by family — Excel's ribbon icons are coloured, not grey line art
// (2026-10-20, "i need the user to be pleased"): layout/totals green, filtering teal, calculations
// indigo, view toggles gold, destructive red, the rest slate.
const ICON_TONES = {
  "Grand Totals": "text-emerald-600",
  Subtotals: "text-emerald-600",
  "Report Layout": "text-emerald-600",
  "Blank Rows": "text-emerald-600",
  "Insert Slicer": "text-teal-600",
  "Insert Timeline": "text-teal-600",
  "Report Filter Pages": "text-sky-600",
  "Calculated Field": "text-indigo-600",
  "Calculated Item": "text-indigo-600",
  "Field List": "text-primary-strong",
  "+/- Buttons": "text-primary-strong",
  "Field Headers": "text-primary-strong",
  Clear: "text-rose-600",
};

function RibbonButtonFace({ icon: Icon, label, badge = 0 }) {
  const compact = useContext(RibbonCompact);
  const iconOnly = useContext(RibbonIconOnly);
  const tone = ICON_TONES[label] ?? "text-slate-600";
  return (
    <>
      <span title={iconOnly ? label : undefined} className="flex shrink-0">
        <Icon className={compact ? `h-4 w-4 shrink-0 ${tone}` : `h-6 w-6 ${tone}`} strokeWidth={1.6} aria-hidden="true" />
      </span>
      <span className={iconOnly ? "sr-only" : compact ? "text-xs leading-none text-slate-700" : "text-center text-xs leading-tight text-slate-700"}>{label}</span>
      {badge > 0 && (
        <span
          className={`flex h-4 w-4 items-center justify-center rounded-full bg-primary text-[10px] font-bold text-white ${compact ? "ml-0.5" : "absolute right-1 top-1"}`}
        >
          {badge}
        </span>
      )}
    </>
  );
}

// A plain (non-dropdown) ribbon button — Options, and the Show group's toggles (`pressed`).
function RibbonPlainButton({ icon, label, onClick, pressed, disabled = false, title }) {
  const compact = useContext(RibbonCompact);
  return (
    <button
      type="button"
      aria-pressed={pressed}
      onClick={onClick}
      disabled={disabled}
      title={title}
      className={ribbonButtonClass(compact, (pressed ? "bg-primary-soft ring-1 ring-inset ring-primary/50" : "hover:bg-slate-100") + " disabled:cursor-not-allowed disabled:opacity-40 disabled:hover:bg-transparent")}
    >
      <RibbonButtonFace icon={icon} label={label} />
    </button>
  );
}

function RibbonButton({ icon, label, badge, panelClassName = "", children }) {
  const compact = useContext(RibbonCompact);
  return (
    <Popover
      panelClassName={`!mt-6 ${panelClassName}`}
      // While its popup is open the button stays marked (the Show toggles' pressed look), like Excel
      // keeps a command highlighted while its gallery or dialog is up (2026-10-25).
      trigger={(toggle, open) => (
        <button
          type="button"
          onClick={toggle}
          aria-expanded={open}
          aria-haspopup="dialog"
          className={ribbonButtonClass(compact, open ? "bg-primary-soft ring-1 ring-inset ring-primary/50" : "hover:bg-slate-100")}
        >
          <RibbonButtonFace icon={icon} label={label} badge={badge} />
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

// Excel's Home › Number group (2026-10-20): the format dropdown over $ (Accounting, with a currency
// menu), % (Percent Style), , (Comma Style) and Increase / Decrease Decimal. Like Excel it formats
// just the selected cells (drag over cells, Shift+click, or click a column header for a column).
const CURRENCY_SYMBOLS = ["", "EGP", "SAR", "AED", "$", "€", "£"];

function NumberGroup({ selection, onApply }) {
  const enabled = Boolean(selection?.count || selection?.cellCount);
  const shared = enabled ? selection.formatType ?? null : null;
  // fn(format, shownDecimals) → the cell's next format; the grid runs it for each selected cell.
  const apply = (fn) => enabled && onApply(fn);
  const hint = enabled ? null : "Select cells in the PivotTable first (click a column header to select a whole column)";
  const iconButton = "flex h-7 min-w-7 items-center justify-center rounded px-1 text-slate-700 hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-40 disabled:hover:bg-transparent";
  return (
    <RibbonGroup label="Number">
      <div className="row-span-3 flex flex-col justify-center gap-1.5 px-1 py-1" title={hint ?? undefined}>
        <select
          aria-label="Number Format"
          disabled={!enabled}
          value={shared ?? ""}
          onChange={(e) => e.target.value && apply((f) => withFormatType(f, e.target.value))}
          className="h-7 w-40 rounded border border-slate-300 bg-white px-1.5 text-xs text-ink focus:outline-none focus:ring-1 focus:ring-primary disabled:cursor-not-allowed disabled:bg-slate-50 disabled:text-muted"
        >
          {shared === null && <option value="">{enabled ? "(mixed)" : "General"}</option>}
          {NUMBER_FORMAT_CHOICES.map((c) => (
            <option key={c.type} value={c.type}>
              {c.label}
            </option>
          ))}
        </select>
        <div className="flex items-center gap-0.5">
          <button type="button" disabled={!enabled} onClick={() => apply((f) => accountingStyle(f))} className={iconButton} title="Accounting Number Format" aria-label="Accounting Number Format">
            <DollarSign className="h-4 w-4 text-emerald-600" />
          </button>
          <Popover
            align="left"
            panelClassName="!min-w-[9rem] !max-w-[9rem] !p-1"
            trigger={(toggle, open) => (
              <button type="button" disabled={!enabled} onClick={toggle} className={`${iconButton} !min-w-4 !px-0 ${open ? "!bg-primary-soft" : ""}`} title="Currency symbol" aria-label="Choose the currency symbol">
                <ChevronDown className="h-3 w-3" />
              </button>
            )}
          >
            {(close) => (
              <div className="flex flex-col text-sm">
                {CURRENCY_SYMBOLS.map((sym) => (
                  <button
                    key={sym || "none"}
                    type="button"
                    onClick={() => {
                      apply((f) => accountingStyle(f, sym));
                      close();
                    }}
                    className="rounded-md px-2 py-1.5 text-left hover:bg-slate-100"
                  >
                    {sym || "None"}
                  </button>
                ))}
              </div>
            )}
          </Popover>
          <button type="button" disabled={!enabled} onClick={() => apply(percentStyle)} className={iconButton} title="Percent Style" aria-label="Percent Style">
            <Percent className="h-4 w-4 text-emerald-600" />
          </button>
          <button type="button" disabled={!enabled} onClick={() => apply(commaStyle)} className={`${iconButton} text-base font-bold leading-none text-emerald-700`} title="Comma Style" aria-label="Comma Style">
            ,
          </button>
          <span className="mx-1 h-5 w-px bg-slate-200" aria-hidden="true" />
          <button type="button" disabled={!enabled} onClick={() => apply((f, shown) => changeDecimals(f, +1, shown))} className={iconButton} title="Increase Decimal" aria-label="Increase Decimal">
            <DecimalsArrowRight className="h-4 w-4 text-sky-600" />
          </button>
          <button type="button" disabled={!enabled} onClick={() => apply((f, shown) => changeDecimals(f, -1, shown))} className={iconButton} title="Decrease Decimal" aria-label="Decrease Decimal">
            <DecimalsArrowLeft className="h-4 w-4 text-sky-600" />
          </button>
        </div>
      </div>
    </RibbonGroup>
  );
}

// One ribbon "group" — a cluster of buttons with a shared bottom caption and a divider on its
// right, exactly Excel's own group convention (e.g. "Calculations", "Layout").
function RibbonGroup({ label, children }) {
  const compact = useContext(RibbonCompact);
  return (
    <div className="flex h-full flex-col items-stretch border-r border-slate-200 px-2 pt-1 last:border-r-0">
      {/* Narrow ribbon: small buttons fill each column top to bottom, three to a column. */}
      <div className={compact ? "grid h-16 flex-1 grid-flow-col grid-rows-3 content-start items-center gap-x-1" : "flex flex-1 items-start gap-1"}>{children}</div>
      <p className="pb-0.5 text-center text-[10px] text-slate-400">{label}</p>
    </div>
  );
}

// pivotOptions/onPivotOptionsChange/onOpenPivotTableOptions: the PivotTable group (name + Options);
// omitted, that group isn't drawn.
// fieldListShown/onToggleFieldList: Excel's "Field List" toggle in the Show group; omitted, no button.
// layout: { reportLayout, repeatItemLabels, fieldLayouts, onReportLayoutChange, onRepeatItemLabelsChange,
// onFieldLayoutsChange } — the Design tab's Report Layout / Blank Rows (2026-10-20); omitted, not drawn.
// numberFormat: { selection (the grid's selected range summary, or null), onApply(fn) — formats the
// selected cells } — Home › Number; omitted, the Home tab has no Number group.
export default function PivotToolbar({ request, onChange, measureOptions, dimensions = [], allDimensions = null, pivotOptions = null, onPivotOptionsChange = null, onOpenPivotTableOptions = null, fieldListShown = true, onToggleFieldList = null, layout = null, actions = null, numberFormat = null, history = null }) {
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

  // Every field of the cube once the full list has loaded (allDimensions), like Excel's Insert Slicers /
  // Insert Timelines — not just the tables expanded in the Field List (2026-10-20).
  const insertable = allDimensions?.length ? allDimensions : dimensions;
  const availableForSlicer = insertable.filter((d) => !request.slicerFields.includes(d.field));
  const availableForTimeline = insertable.filter((d) => d.isDate && !request.timelineFields.includes(d.field));

  // Narrow ribbon (see RibbonCompact): the full-size ribbon's width is measured while it's full size
  // (per tab — each tab has its own groups), and the small buttons are used whenever the space is
  // narrower than that. Measured before paint, so there's no flicker.
  const surfaceRef = useRef(null);
  const [compact, setCompact] = useState(false);
  // Excel's Collapse the Ribbon (Ctrl+F1, 2026-10-20): only the tab names stay, giving the grid the
  // ribbon's height; clicking a tab brings the ribbon back. Remembered per browser.
  const [collapsed, setCollapsed] = useState(() => {
    try {
      return localStorage.getItem(RIBBON_COLLAPSED_KEY) === "true";
    } catch {
      return false;
    }
  });
  useEffect(() => {
    try {
      localStorage.setItem(RIBBON_COLLAPSED_KEY, String(collapsed));
    } catch {
      // Not remembered — still applies on this page.
    }
  }, [collapsed]);
  useEffect(() => {
    const onKey = (e) => {
      if (e.ctrlKey && e.key === "F1") {
        e.preventDefault();
        setCollapsed((v) => !v);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);
  // Each step's width is measured while it's showing (full, then small buttons); the narrowest step
  // that fits is used. Re-checked after every step change, since a step change alters scrollWidth
  // without resizing the element (ResizeObserver alone wouldn't fire).
  const [iconOnly, setIconOnly] = useState(false);
  const compactRef = useRef(false);
  const iconOnlyRef = useRef(false);
  const fullWidthRef = useRef(0);
  const compactWidthRef = useRef(0);
  useLayoutEffect(() => {
    const el = surfaceRef.current;
    if (!el) return undefined;
    const check = () => {
      if (!compactRef.current) fullWidthRef.current = el.scrollWidth;
      else if (!iconOnlyRef.current) compactWidthRef.current = el.scrollWidth;
      const nextCompact = el.clientWidth < fullWidthRef.current;
      const nextIconOnly = nextCompact && compactWidthRef.current > 0 && el.clientWidth < compactWidthRef.current;
      if (nextCompact !== compactRef.current) {
        compactRef.current = nextCompact;
        setCompact(nextCompact);
      }
      if (nextIconOnly !== iconOnlyRef.current) {
        iconOnlyRef.current = nextIconOnly;
        setIconOnly(nextIconOnly);
      }
    };
    check();
    const observer = new ResizeObserver(check);
    observer.observe(el);
    return () => observer.disconnect();
  }, [tab, collapsed, compact, iconOnly]);
  const switchTab = (key) => {
    // Re-measure the new tab at full size first.
    compactRef.current = false;
    iconOnlyRef.current = false;
    compactWidthRef.current = 0;
    setCompact(false);
    setIconOnly(false);
    setTab(key);
    setCollapsed(false);
  };

  return (
    <RibbonCompact.Provider value={compact}>
    <RibbonIconOnly.Provider value={iconOnly}>
    <div className="shrink-0 overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm">
      {/* Tab strip — the modern Office ribbon (2026-10-20 UI refresh): flat, one surface, the active
          tab marked by a short gold underline and bold ink text (was a tinted strip with gold text). */}
      <div role="tablist" aria-label="Ribbon" className={`flex items-center gap-1 px-2 ${collapsed ? "" : "border-b border-slate-200"}`}>
        {TABS.map((t) => {
          const active = tab === t.key && !collapsed;
          return (
            <button
              key={t.key}
              type="button"
              role="tab"
              aria-selected={active}
              onClick={() => switchTab(t.key)}
              className={`relative px-2.5 py-2 text-xs transition-colors ${active ? "font-semibold text-ink" : "font-medium text-slate-500 hover:text-ink"}`}
            >
              {t.label}
              {active && <span className="absolute inset-x-2.5 bottom-0 h-[2px] rounded-full bg-primary" aria-hidden="true" />}
            </button>
          );
        })}
        <button
          type="button"
          onClick={() => setCollapsed((v) => !v)}
          className="ml-auto rounded p-1 text-slate-400 hover:bg-slate-100 hover:text-ink"
          aria-label={collapsed ? "Expand the ribbon" : "Collapse the ribbon"}
          title={collapsed ? "Expand the ribbon (Ctrl+F1)" : "Collapse the ribbon (Ctrl+F1)"}
        >
          <ChevronUp className={`h-4 w-4 transition-transform ${collapsed ? "rotate-180" : ""}`} />
        </button>
      </div>

      {/* Ribbon surface — flat white, groups split by hairlines. */}
      {/* overflow-x-auto (2026-10-19): on a narrow window the groups scroll sideways instead of
          being cut off (the Calculations/Show groups used to vanish). Dropdown panels render in a
          portal, so this doesn't clip them. */}
      {!collapsed && (
      <div ref={surfaceRef} role="tabpanel" className="custom-scrollbar flex items-stretch gap-0 overflow-x-auto bg-white px-1 py-1">
        {/* Excel's Undo group, first on Home (2026-10-21) — Ctrl+Z / Ctrl+Y do the same. */}
        {tab === "home" && history && (
          <RibbonGroup label="Undo">
            <RibbonPlainButton icon={Undo2} label="Undo" onClick={history.undo} disabled={!history.canUndo} title="Undo (Ctrl+Z)" />
            <RibbonPlainButton icon={Redo2} label="Redo" onClick={history.redo} disabled={!history.canRedo} title="Redo (Ctrl+Y)" />
          </RibbonGroup>
        )}
        {tab === "home" && numberFormat && <NumberGroup {...numberFormat} />}
        {tab === "home" && !numberFormat && <p className="px-3 py-6 text-xs text-muted">Nothing here yet.</p>}

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
                      className={`rounded-lg px-2 py-1.5 text-left ${request.grandTotals === opt.value ? "bg-primary-soft font-semibold text-primary-strong" : "hover:bg-slate-100"}`}
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
                          request.subtotals === opt.value ? "bg-primary-soft font-semibold text-primary-strong" : "hover:bg-slate-100"
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
                          className={`rounded-lg px-2 py-1.5 text-left ${layout.reportLayout === opt.value ? "bg-primary-soft font-semibold text-primary-strong" : "hover:bg-slate-100"}`}
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
                            layout.reportLayout !== "Compact" && layout.repeatItemLabels === opt.value ? "bg-primary-soft font-semibold text-primary-strong" : "hover:bg-slate-100"
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
                          className="rounded-md px-2 py-1.5 text-left hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-40"
                        >
                          Insert Blank Line after Each Item
                        </button>
                        <button type="button" onClick={() => setBlankLines(false)} className="rounded-md px-2 py-1.5 text-left hover:bg-slate-100">
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
            <div className={`flex flex-col justify-center gap-1 px-1 ${compact ? "row-span-3 h-16 w-32" : "h-16 w-40"}`}>
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
                className="rounded border border-slate-300 bg-white px-1.5 py-0.5 text-xs focus:outline-none focus:ring-1 focus:ring-primary"
              />
            </div>
            <RibbonPlainButton icon={Settings2} label="Options" onClick={onOpenPivotTableOptions} />
            {/* Excel's Options ▸ Show Report Filter Pages (2026-10-20): one copy of the pivot per item
                of a Report Filter field — here, one sheet each in a downloaded workbook. */}
            {actions && (
              <RibbonButton icon={Files} label="Report Filter Pages" badge={0}>
                {(close) => (
                  <div className="w-full space-y-2 text-sm">
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
                              className="rounded-md px-2 py-1.5 text-left hover:bg-slate-100"
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
                    className="rounded-md px-2 py-1.5 text-left hover:bg-slate-100"
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
                    className="rounded-md px-2 py-1.5 text-left hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-40"
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
            {onToggleFieldList && <RibbonPlainButton icon={PanelRight} label="Field List" pressed={fieldListShown} onClick={onToggleFieldList} />}
            {[
              { key: "showExpandButtons", label: "+/- Buttons", icon: SquarePlus },
              { key: "showFieldCaptions", label: "Field Headers", icon: Heading },
            ].map(({ key, label, icon }) => {
              const on = pivotOptions[key] !== false;
              return <RibbonPlainButton key={key} icon={icon} label={label} pressed={on} onClick={() => onPivotOptionsChange({ ...pivotOptions, [key]: !on })} />;
            })}
          </RibbonGroup>
        )}
      </div>
      )}
    </div>
    </RibbonIconOnly.Provider>
    </RibbonCompact.Provider>
  );
}
