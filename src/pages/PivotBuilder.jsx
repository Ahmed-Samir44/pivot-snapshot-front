import { useEffect, useState } from "react";
import { FileSpreadsheet } from "lucide-react";
import AppShell from "../components/layout/AppShell";
import FieldPicker from "../components/pivot/FieldPicker";
import PivotGrid from "../components/pivot/PivotGrid";
import PivotChart from "../components/pivot/PivotChart";
import { chartUnavailableReason } from "../utils/chartLayout";
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
  rowLevelSorts: [],
  columnLevelSorts: [],
  showGrandTotals: false,
  showSubtotals: false,
  calculatedFields: [],
  calculatedItems: [],
  dateGroupings: [],
  numericGroupings: [],
  showItemsWithNoData: [],
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
  // Set only for SnapshotTooLargeException ("SnapshotTooLarge") — mirrors cellsWarning's "warn,
  // then let the user explicitly confirm" flow, but (like cardinalityWarning, unlike cellsWarning)
  // this one has a hard ceiling: a saved snapshot is reopened by other people later, not a
  // one-time view only the current user sees (see StorageOptions.AbsoluteMaxSnapshotRows).
  const [rowsWarning, setRowsWarning] = useState(null);
  const [exporting, setExporting] = useState(false);
  const [exportError, setExportError] = useState(null);

  // Set only for CardinalityGuard's HighCardinalityFieldException (see pivotApi.js's
  // error.details) — a distinct state from `error` because this one gets its own warning UI with
  // an explicit "run anyway" retry, not just an error message.
  const [cardinalityWarning, setCardinalityWarning] = useState(null);

  // Set only for PivotResultTooLargeException ("ResultTooLarge") — unlike cardinalityWarning,
  // there's no absolute ceiling here (see PivotRequest.MaxCellsOverride's WHY comment): this is
  // the browser's own "this table will be huge" warning, not a shared-server crash risk, so the
  // "run anyway" retry always works once the user confirms.
  const [cellsWarning, setCellsWarning] = useState(null);
  // Sticky accumulator for maxMembersOverride/maxCellsOverride once the user explicitly confirms
  // one of the two warnings below — see runQuery's own WHY comment for why this must NOT reset
  // per-call (a real ping-ponging-warnings bug caught live 2026-09-27 when it did).
  const [confirmedOverrides, setConfirmedOverrides] = useState({});

  useEffect(() => {
    Promise.all([getDimensions(), getMeasures()])
      .then(([dims, meas]) => {
        setDimensions(dims);
        setMeasures(meas);
      })
      .catch((err) => setMetaError(err.message));
  }, []);

  // newOverride is only ever set by the user explicitly confirming one of the two warnings below
  // (runQueryWithOverride/runQueryWithCellsOverride) — never sent by default, so a normal "Run
  // query" click always goes through the ordinary safe limits. It's merged with confirmedOverrides
  // (not replacing it): the backend checks cardinality first, then result size only once that
  // passes, so overriding cardinality can surface a FRESH "result too large" warning next — and
  // confirming that one must still carry the already-confirmed cardinality override forward, or
  // the very next attempt fails the cardinality check all over again.
  const runQuery = async (newOverride = {}) => {
    setLoading(true);
    setError(null);
    setCardinalityWarning(null);
    setCellsWarning(null);
    setSaveMessage(null);
    const overrides = { ...confirmedOverrides, ...newOverride };
    // Persisted BEFORE the request even runs, not just on success: confirming maxMembersOverride
    // can get an attempt PAST the cardinality check but still fail later at the result-size check
    // (a fresh, different warning) — that later failure must not un-remember the override that
    // already worked, or the very next retry loses it and fails cardinality all over again. This
    // was the actual remaining half of the 2026-09-27 ping-ponging-warnings bug: the first fix
    // (merging confirmedOverrides into every attempt) was correct but only PERSISTED the merge on
    // a fully successful query, so a retry that hit a SECOND, later warning still lost the first
    // one's confirmation.
    setConfirmedOverrides(overrides);
    try {
      const cleaned = {
        rows: request.rows.filter(Boolean),
        columns: request.columns.filter(Boolean),
        values: request.values.filter((v) => v.field),
        valuesPlacement: request.valuesPlacement,
        filters: request.filters.filter((f) => f.field && isFilterComplete(f)),
        sort: request.sort,
        rowLevelSorts: request.rowLevelSorts,
        columnLevelSorts: request.columnLevelSorts,
        showGrandTotals: request.showGrandTotals,
        showSubtotals: request.showSubtotals,
        calculatedFields: request.calculatedFields.filter((f) => f.name && f.leftField && f.rightField),
        // A CalculatedItem needs a field, a name, and at least one member on either side (see
        // MdxPivotQueryBuilder.ValidateCalculatedItems) — mirrors the calculatedFields cleaning
        // above.
        calculatedItems: request.calculatedItems.filter(
          (i) => i.field && i.name && (i.positiveMembers.length > 0 || i.negativeMembers.length > 0),
        ),
        // A grouped field can be in Rows OR Columns (see GroupingProcessor) — checking only Rows
        // here would silently strip a Columns-axis grouping right before it's sent.
        dateGroupings: request.dateGroupings.filter((g) => request.rows.includes(g.field) || request.columns.includes(g.field)),
        numericGroupings: request.numericGroupings.filter(
          (g) => (request.rows.includes(g.field) || request.columns.includes(g.field)) && g.binSize > 0,
        ),
        showItemsWithNoData: request.showItemsWithNoData,
        ...(overrides.maxMembersOverride ? { maxMembersOverride: overrides.maxMembersOverride } : {}),
        ...(overrides.maxCellsOverride ? { maxCellsOverride: overrides.maxCellsOverride } : {}),
      };
      const data = await queryPivot(cleaned);
      setResult(data);
      setLastQuery(cleaned);
    } catch (err) {
      if (err.details?.type === "HighCardinalityField") {
        setCardinalityWarning(err.details);
      } else if (err.details?.type === "ResultTooLarge") {
        setCellsWarning(err.details);
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
      runQuery({ maxMembersOverride: cardinalityWarning.absoluteMaxMembersPerField });
    }
  };

  const runQueryWithCellsOverride = () => {
    if (cellsWarning) {
      runQuery({ maxCellsOverride: cellsWarning.actualCells });
    }
  };

  const handleSave = async (maxRowsOverride) => {
    if (!snapshotName.trim() || !result || !lastQuery) {
      return;
    }
    setSaving(true);
    setSaveError(null);
    setSaveMessage(null);
    setRowsWarning(null);
    try {
      const saved = await saveSnapshot(snapshotName.trim(), lastQuery, result, tableStyle, includeChart, chartType, reportLayout, maxRowsOverride);
      setSaveMessage(`Saved as version ${saved.versionNumber}.`);
    } catch (err) {
      if (err.details?.type === "SnapshotTooLarge") {
        setRowsWarning(err.details);
      } else {
        setSaveError(err.message);
      }
    } finally {
      setSaving(false);
    }
  };

  const handleSaveWithRowsOverride = () => {
    if (rowsWarning) {
      handleSave(rowsWarning.actualRows);
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

      {cellsWarning && (
        <div className="mt-4 rounded-xl border border-amber-300 bg-amber-50 px-4 py-3 text-amber-900">
          <p>{cellsWarning.error}</p>
          <p className="mt-1 text-sm">
            A table this large can be slow to scroll or read — unlike the cube-side limit above, there&apos;s no hard ceiling here, so you can proceed if
            you&apos;re sure this is what you want.
          </p>
          <button type="button" onClick={runQueryWithCellsOverride} disabled={loading} className="btn-secondary mt-2 disabled:cursor-not-allowed disabled:opacity-50">
            Show all {cellsWarning.actualCells.toLocaleString()} cells anyway
          </button>
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

      {includeChart &&
        chartUnavailableReason(result, lastQuery) && (
          <p className="mt-4 rounded-xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm text-muted">{chartUnavailableReason(result, lastQuery)}</p>
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
            <option value="Outline">Layout: Outline</option>
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
            onClick={() => handleSave()}
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

      {rowsWarning && (
        <div className="mt-4 rounded-xl border border-amber-300 bg-amber-50 px-4 py-3 text-amber-900">
          <p>{rowsWarning.error}</p>
          {rowsWarning.canOverride ? (
            <>
              <p className="mt-1 text-sm">
                A saved snapshot is reopened by other people later, so an unusually large one stays slow/heavy for everyone who opens it — unlike the
                live view above, this has a hard ceiling. Only proceed if you&apos;re sure this size is needed.
              </p>
              <button type="button" onClick={handleSaveWithRowsOverride} disabled={saving} className="btn-secondary mt-2 disabled:cursor-not-allowed disabled:opacity-50">
                I understand, save all {rowsWarning.actualRows.toLocaleString()} rows anyway
              </button>
            </>
          ) : (
            <p className="mt-1 text-sm">
              The row count ({rowsWarning.actualRows.toLocaleString()}) exceeds the absolute limit ({rowsWarning.absoluteMaxRows.toLocaleString()}) — you
              must narrow the pivot before saving; there is no override available here.
            </p>
          )}
        </div>
      )}
    </AppShell>
  );
}
