import { useEffect, useState } from "react";
import { FileSpreadsheet } from "lucide-react";
import AppShell from "../components/layout/AppShell";
import FieldPicker from "../components/pivot/FieldPicker";
import PivotGrid from "../components/pivot/PivotGrid";
import PivotChart from "../components/pivot/PivotChart";
import { queryPivot } from "../services/pivotApi";
import { getDimensions, getMeasures } from "../services/cubeMetaApi";
import { saveSnapshot } from "../services/snapshotApi";

// This screen renders a live, interactive pivot (React state + re-fetch on every change).
// It is NOT the saved snapshot: "Save Snapshot" below freezes the CURRENT result (the one already
// on screen) into a static HTML snapshot — it never re-queries the cube, so what gets saved is
// guaranteed to match exactly what the user looked at before saving (see SnapshotStorageService).
const EMPTY_REQUEST = {
  rows: [],
  columns: [],
  values: [],
  valuesPlacement: "Columns",
  filters: [],
  sort: null,
  showGrandTotals: false,
  showSubtotals: false,
  calculatedFields: [],
  calculatedItems: [],
  dateGroupings: [],
  numericGroupings: [],
};

// PivotResult only carries member VALUES (e.g. "Cardiology"), not the field's own name — this
// looks each row/column field id up in the dimensions list to get something readable ("Specialty
// Name") for the grid's header row, instead of leaving it blank.
function fieldLabels(fieldIds, dimensions) {
  if (!fieldIds) return [];
  return fieldIds.map((id) => dimensions.find((d) => d.field === id)?.displayName ?? id);
}

// A filter is only ready to send once whichever fields ITS mode actually needs are filled in —
// each mode uses a different subset of PivotFilter's optional fields (see the backend's
// MdxPivotQueryBuilder.ValidateFilters for the authoritative per-mode rules this mirrors).
function isFilterComplete(filter) {
  switch (filter.mode ?? "Members") {
    case "LabelContains":
    case "LabelBeginsWith":
    case "LabelEndsWith":
      return Boolean(filter.labelText);
    case "TopN":
    case "BottomN":
      return Boolean(filter.n > 0 && filter.byMeasureField);
    default:
      return filter.includedMembers.length > 0;
  }
}

export default function PivotBuilder() {
  const [request, setRequest] = useState(EMPTY_REQUEST);
  const [lastQuery, setLastQuery] = useState(null); // the exact cleaned request that produced `result`
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [dimensions, setDimensions] = useState([]);
  const [measures, setMeasures] = useState([]);
  const [metaError, setMetaError] = useState(null);

  const [snapshotName, setSnapshotName] = useState("");
  const [tableStyle, setTableStyle] = useState("Default");
  const [includeChart, setIncludeChart] = useState(false);
  const [chartType, setChartType] = useState("Bar");
  const [reportLayout, setReportLayout] = useState("Tabular");
  const [saving, setSaving] = useState(false);
  const [saveMessage, setSaveMessage] = useState(null);
  const [saveError, setSaveError] = useState(null);
  const [exporting, setExporting] = useState(false);
  const [exportError, setExportError] = useState(null);

  // Set only for CardinalityGuard's HighCardinalityFieldException (see pivotApi.js's
  // error.details) — a distinct state from `error` because this one gets its own warning UI with
  // an explicit "run anyway" retry, not just an error message.
  const [cardinalityWarning, setCardinalityWarning] = useState(null);

  useEffect(() => {
    Promise.all([getDimensions(), getMeasures()])
      .then(([dims, meas]) => {
        setDimensions(dims);
        setMeasures(meas);
      })
      .catch((err) => setMetaError(err.message));
  }, []);

  // maxMembersOverride is only ever set by the user explicitly confirming the cardinality
  // warning below (runQueryWithOverride) — never sent by default, so a normal "Run query" click
  // always goes through CardinalityGuard's ordinary safe limit.
  const runQuery = async (maxMembersOverride) => {
    setLoading(true);
    setError(null);
    setCardinalityWarning(null);
    setSaveMessage(null);
    try {
      const cleaned = {
        rows: request.rows.filter(Boolean),
        columns: request.columns.filter(Boolean),
        values: request.values.filter((v) => v.field),
        valuesPlacement: request.valuesPlacement,
        filters: request.filters.filter((f) => f.field && isFilterComplete(f)),
        sort: request.sort,
        showGrandTotals: request.showGrandTotals,
        showSubtotals: request.showSubtotals,
        calculatedFields: request.calculatedFields.filter((f) => f.name && f.leftField && f.rightField),
        // A CalculatedItem needs a field, a name, and at least one member on either side (see
        // MdxPivotQueryBuilder.ValidateCalculatedItems) — mirrors the calculatedFields cleaning
        // above.
        calculatedItems: request.calculatedItems.filter(
          (i) => i.field && i.name && (i.positiveMembers.length > 0 || i.negativeMembers.length > 0),
        ),
        dateGroupings: request.dateGroupings.filter((g) => request.rows.includes(g.field)),
        numericGroupings: request.numericGroupings.filter((g) => request.rows.includes(g.field) && g.binSize > 0),
        ...(maxMembersOverride ? { maxMembersOverride } : {}),
      };
      const data = await queryPivot(cleaned);
      setResult(data);
      setLastQuery(cleaned);
    } catch (err) {
      if (err.details?.type === "HighCardinalityField") {
        setCardinalityWarning(err.details);
      } else {
        setError(err.message);
      }
      setResult(null);
      setLastQuery(null);
    } finally {
      setLoading(false);
    }
  };

  const runQueryWithOverride = () => {
    if (cardinalityWarning) {
      runQuery(cardinalityWarning.absoluteMaxMembersPerField);
    }
  };

  const handleSave = async () => {
    if (!snapshotName.trim() || !result || !lastQuery) {
      return;
    }
    setSaving(true);
    setSaveError(null);
    setSaveMessage(null);
    try {
      const saved = await saveSnapshot(snapshotName.trim(), lastQuery, result, tableStyle, includeChart, chartType, reportLayout);
      setSaveMessage(`Saved as version ${saved.versionNumber}.`);
    } catch (err) {
      setSaveError(err.message);
    } finally {
      setSaving(false);
    }
  };

  const handleExport = async () => {
    if (!result || !lastQuery) return;
    setExporting(true);
    setExportError(null);
    try {
      // Dynamic import: ExcelJS is a genuinely large library (bundle jumped from ~420KB to
      // ~1.35MB when statically imported, 2026-09-23) — loading it only when someone actually
      // clicks Export keeps it out of everyone else's initial page load.
      const { exportPivotToExcel } = await import("../utils/exportExcel");
      await exportPivotToExcel({
        result,
        rowFieldLabels: fieldLabels(lastQuery.rows, dimensions),
        columnFieldLabels: fieldLabels(lastQuery.columns, dimensions),
        valueFields: lastQuery.values,
        valuesPlacement: lastQuery.valuesPlacement,
        filename: snapshotName.trim() || "Pivot",
      });
    } catch (err) {
      setExportError(err.message);
    } finally {
      setExporting(false);
    }
  };

  return (
    <AppShell>
      <header className="mb-8">
        <h1 className="text-3xl font-bold text-ink">Pivot Snapshot Builder</h1>
        <p className="mt-1 text-muted">Build a pivot, then save it as a snapshot.</p>
      </header>

      {metaError && (
        <p className="mb-6 rounded-xl border border-amber-300 bg-amber-50 px-4 py-3 text-amber-800">
          Couldn&apos;t load the field list from the API: {metaError}
        </p>
      )}

      <div className="mb-6">
        <FieldPicker value={request} onChange={setRequest} dimensions={dimensions} measures={measures} />
      </div>

      <button type="button" onClick={() => runQuery()} disabled={loading} className="btn-primary disabled:cursor-not-allowed disabled:opacity-50">
        {loading ? "Loading…" : "Run query"}
      </button>

      {error && <p className="mt-4 rounded-xl border border-red-300 bg-red-50 px-4 py-3 text-red-800">{error}</p>}

      {cardinalityWarning && (
        <div className="mt-4 rounded-xl border border-amber-300 bg-amber-50 px-4 py-3 text-amber-900">
          <p>{cardinalityWarning.error}</p>
          {cardinalityWarning.canOverride ? (
            <>
              <p className="mt-1 text-sm">
                This server is shared with other projects and has crashed before from a similar query — running this without a filter could slow down or take
                down the server for other users. Only proceed if you&apos;re sure.
              </p>
              <button type="button" onClick={runQueryWithOverride} disabled={loading} className="btn-secondary mt-2 disabled:cursor-not-allowed disabled:opacity-50">
                I understand the risk, run anyway (up to {cardinalityWarning.absoluteMaxMembersPerField.toLocaleString()} members)
              </button>
            </>
          ) : (
            <p className="mt-1 text-sm">
              The member count ({cardinalityWarning.memberCount.toLocaleString()}) exceeds the server&apos;s absolute limit (
              {cardinalityWarning.absoluteMaxMembersPerField.toLocaleString()}) — you must add a filter on this field; there is no override available here.
            </p>
          )}
        </div>
      )}

      <div className="mt-6">
        <PivotGrid
          result={result}
          rowFieldLabels={fieldLabels(lastQuery?.rows, dimensions)}
          columnFieldLabels={fieldLabels(lastQuery?.columns, dimensions)}
          valueFields={lastQuery?.values}
          layout={reportLayout}
          valuesPlacement={lastQuery?.valuesPlacement}
        />
      </div>

      {includeChart && lastQuery?.valuesPlacement === "Rows" && (
        <p className="mt-4 text-sm text-muted">
          No chart is drawn while values are on Rows — see the "Show values on" setting above.
        </p>
      )}
      {includeChart && (
        <PivotChart result={result} valueFields={lastQuery?.values} chartType={chartType} valuesPlacement={lastQuery?.valuesPlacement} />
      )}

      {result && (
        <div className="card mt-6 flex flex-wrap items-center gap-3">
          <input
            type="text"
            className="input-field max-w-xs"
            placeholder="Snapshot name…"
            value={snapshotName}
            onChange={(e) => setSnapshotName(e.target.value)}
          />
          <select
            className="rounded-xl border border-slate-200 bg-white/90 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent"
            value={tableStyle}
            onChange={(e) => setTableStyle(e.target.value)}
            title="Table style for the saved snapshot"
          >
            <option value="Default">Style: Default</option>
            <option value="Banded">Style: Banded rows</option>
            <option value="Dark">Style: Dark</option>
            <option value="Minimal">Style: Minimal</option>
          </select>
          <select
            className="rounded-xl border border-slate-200 bg-white/90 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent"
            value={reportLayout}
            onChange={(e) => setReportLayout(e.target.value)}
            title="Report layout for the Rows area"
          >
            <option value="Tabular">Layout: Tabular</option>
            <option value="Compact">Layout: Compact</option>
          </select>
          <label className="flex items-center gap-2 text-sm text-ink">
            <input type="checkbox" checked={includeChart} onChange={(e) => setIncludeChart(e.target.checked)} />
            Include chart
          </label>
          {includeChart && (
            <select
              className="rounded-xl border border-slate-200 bg-white/90 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent"
              value={chartType}
              onChange={(e) => setChartType(e.target.value)}
              title="Chart type"
            >
              <option value="Bar">Chart: Bar</option>
              <option value="Line">Chart: Line</option>
              <option value="Pie">Chart: Pie</option>
            </select>
          )}
          <button
            type="button"
            onClick={handleSave}
            disabled={saving || !snapshotName.trim()}
            className="btn-primary disabled:cursor-not-allowed disabled:opacity-50"
          >
            {saving ? "Saving…" : "Save Snapshot"}
          </button>
          <button
            type="button"
            onClick={handleExport}
            disabled={exporting}
            className="btn-secondary inline-flex items-center gap-1.5 disabled:cursor-not-allowed disabled:opacity-50"
          >
            <FileSpreadsheet className="h-4 w-4" style={{ color: "#217346" }} />
            {exporting ? "Exporting…" : "Export to Excel"}
          </button>
          {saveMessage && <span className="text-sm font-medium text-green-700">{saveMessage}</span>}
          {saveError && <span className="text-sm font-medium text-red-700">{saveError}</span>}
          {exportError && <span className="text-sm font-medium text-red-700">{exportError}</span>}
        </div>
      )}
    </AppShell>
  );
}
