import { useState } from "react";
import Modal from "../forms/Modal";
import { DEFAULT_PIVOT_OPTIONS } from "../../utils/pivotOptions";

const selectClass =
  "rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-primary focus:border-transparent";
const textClass =
  "rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-primary focus:border-transparent disabled:bg-slate-100 disabled:text-muted";

const TABS = [
  { key: "layout", label: "Layout & Format" },
  { key: "totals", label: "Totals & Filters" },
  { key: "display", label: "Display" },
  { key: "printing", label: "Printing" },
  { key: "data", label: "Data" },
  { key: "altText", label: "Alt Text" },
];

function Section({ title, children }) {
  return (
    <fieldset className="space-y-2">
      <legend className="mb-1 w-full border-b border-slate-200 pb-1 text-xs font-semibold text-ink">{title}</legend>
      {children}
    </fieldset>
  );
}

function Check({ label, checked, onChange, children }) {
  return (
    <label className="flex items-center gap-2 text-sm text-ink">
      <input type="checkbox" className="h-4 w-4 accent-primary" checked={checked} onChange={(e) => onChange(e.target.checked)} />
      <span className="min-w-[10rem]">{label}</span>
      {children}
    </label>
  );
}

// Excel's PivotTable Options dialog (2026-10-18, against a screenshot of the real one): the name on
// top, then tabs. Edits a DRAFT and applies it on OK, like Excel. Layout & Format holds the display
// settings (`options`, beside `request` — see utils/pivotOptions.js); Totals & Filters holds the
// Grand Totals switches, which live on `request` (they change the query result).
// Not built (not applicable or no screenshot yet): "When in compact form indent row labels" (the
// grid uses a fixed indent), "Autofit column widths on update" (the grid always fits its content),
// "Preserve cell formatting on update" (formatting is stored on the report, so it's always kept).
export default function PivotTableOptionsDialog({ options, request, onApply, onClose }) {
  const [tab, setTab] = useState("layout");
  const [draft, setDraft] = useState({ ...DEFAULT_PIVOT_OPTIONS, ...options });
  const [grandTotals, setGrandTotals] = useState(request.grandTotals);
  const [includeFilteredItems, setIncludeFilteredItems] = useState(Boolean(request.includeFilteredItemsInTotals));
  const [allowMultipleFilters, setAllowMultipleFilters] = useState(Boolean(request.allowMultipleFiltersPerField));
  // Display tab's "Show items with no data on rows/columns" — the same per-axis setting as each
  // field menu's "Show items with no data" (request.showItemsWithNoData).
  const [noDataAxes, setNoDataAxes] = useState(request.showItemsWithNoData ?? []);
  const toggleNoData = (axis, on) => setNoDataAxes((prev) => (on ? [...prev.filter((a) => a !== axis), axis] : prev.filter((a) => a !== axis)));
  const set = (patch) => setDraft((prev) => ({ ...prev, ...patch }));

  // Excel's two checkboxes over this app's 4-way GrandTotals (the Design ribbon's dropdown).
  const rowsOn = grandTotals === "RowsAndColumns" || grandTotals === "RowsOnly";
  const columnsOn = grandTotals === "RowsAndColumns" || grandTotals === "ColumnsOnly";
  const setTotals = (rows, columns) =>
    setGrandTotals(rows && columns ? "RowsAndColumns" : rows ? "RowsOnly" : columns ? "ColumnsOnly" : "Off");

  const apply = () => {
    onApply({ ...draft, name: draft.name.trim() || DEFAULT_PIVOT_OPTIONS.name, filterFieldsPerColumn: Math.max(0, Math.floor(Number(draft.filterFieldsPerColumn) || 0)) }, { grandTotals, includeFilteredItemsInTotals: includeFilteredItems, allowMultipleFiltersPerField: allowMultipleFilters, showItemsWithNoData: noDataAxes });
    onClose();
  };

  return (
    <Modal title="PivotTable Options" onClose={onClose}>
      <div className="w-[30rem] max-w-full space-y-4">
        <label className="flex items-center gap-2 text-sm text-ink">
          <span className="shrink-0">PivotTable Name:</span>
          <input type="text" className={`${textClass} flex-1`} value={draft.name} onChange={(e) => set({ name: e.target.value })} />
        </label>

        <div className="flex gap-1 border-b border-slate-200">
          {TABS.map((t) => (
            <button
              key={t.key}
              type="button"
              onClick={() => setTab(t.key)}
              className={`-mb-px rounded-t-lg border px-3 py-1.5 text-sm ${
                tab === t.key ? "border-slate-200 border-b-white bg-white font-semibold text-ink" : "border-transparent text-muted hover:text-ink"
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>

        {tab === "layout" ? (
          <div className="space-y-4">
            <Section title="Layout">
              <Check label="Merge and center cells with labels" checked={draft.mergeLabels} onChange={(v) => set({ mergeLabels: v })} />
              <label className="flex items-center gap-2 text-sm text-ink">
                <span className="min-w-[14rem]">Display fields in report filter area:</span>
                <select className={selectClass} value={draft.filterAreaOrder} onChange={(e) => set({ filterAreaOrder: e.target.value })}>
                  <option value="DownThenOver">Down, Then Over</option>
                  <option value="OverThenDown">Over, Then Down</option>
                </select>
              </label>
              <label className="flex items-center gap-2 text-sm text-ink">
                <span className="min-w-[14rem]">{draft.filterAreaOrder === "OverThenDown" ? "Report filter fields per row:" : "Report filter fields per column:"}</span>
                <input
                  type="number"
                  min={0}
                  max={255}
                  className={`${textClass} w-20`}
                  value={draft.filterFieldsPerColumn}
                  onChange={(e) => set({ filterFieldsPerColumn: e.target.value })}
                />
              </label>
            </Section>
            <Section title="Format">
              <Check label="For error values show:" checked={draft.showErrorText} onChange={(v) => set({ showErrorText: v })}>
                <input type="text" className={`${textClass} w-32`} disabled={!draft.showErrorText} value={draft.errorText} onChange={(e) => set({ errorText: e.target.value })} />
              </Check>
              <Check label="For empty cells show:" checked={draft.showEmptyText} onChange={(v) => set({ showEmptyText: v })}>
                <input type="text" className={`${textClass} w-32`} disabled={!draft.showEmptyText} value={draft.emptyText} onChange={(e) => set({ emptyText: e.target.value })} />
              </Check>
            </Section>
          </div>
        ) : tab === "totals" ? (
          <div className="space-y-4">
            <Section title="Grand Totals">
              <Check label="Show grand totals for rows" checked={rowsOn} onChange={(v) => setTotals(v, columnsOn)} />
              <Check label="Show grand totals for columns" checked={columnsOn} onChange={(v) => setTotals(rowsOn, v)} />
            </Section>
            {/* Excel's Filters section for a cube report. Off: a total is the sum of the visible
                items; on: it counts a Rows/Columns field's filtered-out items too. */}
            <Section title="Filters">
              <Check label="Include filtered items in totals" checked={includeFilteredItems} onChange={setIncludeFilteredItems} />
              {/* One filter of each kind (manual / label / value) per field at once, applied in
                  that order — off, a new filter on a field replaces its old one. */}
              <Check label="Allow multiple filters per field" checked={allowMultipleFilters} onChange={setAllowMultipleFilters} />
            </Section>
          </div>
        ) : tab === "display" ? (
          // Excel's Display tab, for a cube report. Its other options (tooltips, classic layout,
          // the Values row, item labels with no value fields, field list order) aren't built yet.
          <div className="space-y-4">
            <Section title="Display">
              <Check label="Show expand/collapse buttons" checked={draft.showExpandButtons !== false} onChange={(v) => set({ showExpandButtons: v })} />
              <Check label="Display field captions and filter drop downs" checked={draft.showFieldCaptions !== false} onChange={(v) => set({ showFieldCaptions: v })} />
              <Check label="Show items with no data on rows" checked={noDataAxes.includes("Rows")} onChange={(v) => toggleNoData("Rows", v)} />
              <Check label="Show items with no data on columns" checked={noDataAxes.includes("Columns")} onChange={(v) => toggleNoData("Columns", v)} />
            </Section>
          </div>
        ) : tab === "printing" ? (
          // Excel's Printing tab (2026-10-20) — applied to the Excel export's page setup, since that's
          // what gets printed. ("Print expand/collapse buttons" has nothing to act on: the export has
          // no buttons.)
          <Section title="Print">
            <Check label="Repeat row labels on each printed page" checked={draft.repeatRowLabelsOnPrint !== false} onChange={(v) => set({ repeatRowLabelsOnPrint: v })} />
            <Check label="Set print titles" checked={Boolean(draft.printTitles)} onChange={(v) => set({ printTitles: v })} />
            <p className="text-xs text-muted">Applies to Export to Excel: the row label columns and the header rows repeat on every printed page.</p>
          </Section>
        ) : tab === "altText" ? (
          // Excel's Alt Text tab: a title and description for screen readers — the table's
          // accessible name here, and the exported file's Title/Subject.
          <div className="space-y-3">
            <label className="flex flex-col gap-1 text-sm text-ink">
              Title
              <input type="text" className={textClass} value={draft.altTextTitle ?? ""} onChange={(e) => set({ altTextTitle: e.target.value })} />
            </label>
            <label className="flex flex-col gap-1 text-sm text-ink">
              Description
              <textarea rows={4} className={textClass} value={draft.altTextDescription ?? ""} onChange={(e) => set({ altTextDescription: e.target.value })} />
            </label>
          </div>
        ) : (
          // Excel's Data tab. Its other options (save source data with the file, refresh on open,
          // items to retain per field) are about a local data cache — this report always queries the
          // live cube, so they don't apply.
          <Section title="PivotTable Data">
            <Check label="Enable show details" checked={draft.enableShowDetails !== false} onChange={(v) => set({ enableShowDetails: v })} />
          </Section>
        )}

        <div className="flex justify-end gap-2 pt-2">
          <button type="button" className="btn-primary px-5 py-1.5" onClick={apply}>
            OK
          </button>
          <button type="button" className="btn-secondary px-5 py-1.5" onClick={onClose}>
            Cancel
          </button>
        </div>
      </div>
    </Modal>
  );
}
