import { useEffect, useState } from "react";
import { FileSpreadsheet } from "lucide-react";
import AppShell from "../components/layout/AppShell";
import FieldPicker from "../components/pivot/FieldPicker";
import PivotGrid from "../components/pivot/PivotGrid";
import PivotChart from "../components/pivot/PivotChart";
import { chartUnavailableReason } from "../utils/chartLayout";
import { queryPivot } from "../services/pivotApi";
import { getTables, getDimensionsForTable, getMeasures } from "../services/cubeMetaApi";
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

// Which tables the "Available fields" picker actually works with, out of everything the cube's
// schema has (302 tables as of 2026-09-27, mostly internal/system tables no business user needs).
// Kept in the browser, not the backend: the user asked for a per-user pick ("خلي اليوزر يختار"),
// not one shared setting for everyone.
const ENABLED_TABLES_STORAGE_KEY = "pivotSnapshot.enabledTables";

function readEnabledTables() {
  try {
    const raw = localStorage.getItem(ENABLED_TABLES_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
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

// Everything in `request` EXCEPT each value's `format`/`conditionalFormat` — those two are pure
// display settings (see displayValues' own WHY comment below) that must NOT mark the on-screen
// result as stale, unlike every other field here (rows, columns, filters, sort, groupings,
// totals, showItemsWithNoData, valuesPlacement, calculatedFields/Items, and even a value's
// showValuesAs — that one DOES need a fresh query since it's computed server-side). Used to detect
// "the user changed something that actually invalidates the table currently on screen."
function structuralSignature(req) {
  return JSON.stringify({
    ...req,
    values: req.values.map(({ format: _format, conditionalFormat: _conditionalFormat, ...structural }) => structural),
  });
}

export default function PivotBuilder() {
  const [request, setRequest] = useState(EMPTY_REQUEST);
  const [lastQuery, setLastQuery] = useState(null); // the exact cleaned request that produced `result`
  const [result, setResult] = useState(null);
  // structuralSignature(request) as of the last successful run — compared on every `request` change
  // below to clear a stale result instead of leaving a table on screen that no longer matches the
  // current Rows/Columns/Filters/etc (explicitly requested, 2026-09-27: a table left over from a
  // previous run was confusing once the fields it was built from had already changed).
  const [lastRunSignature, setLastRunSignature] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [measures, setMeasures] = useState([]);
  const [metaError, setMetaError] = useState(null);
  // Separate from `loading` above (that one's for Run query). `tables` is the fast, columns-free
  // catalog (302 tables) the multiselect below picks from; `fieldsByTable` is a cache of each
  // table's own columns ever fetched, fetched lazily one table at a time instead of eagerly
  // loading every table's columns up front (2301 fields, several seconds — explicitly asked for
  // live, 2026-09-27, after the earlier "load everything" version made the panel sit empty that
  // whole time with nothing saying so).
  const [tables, setTables] = useState([]);
  const [tablesLoading, setTablesLoading] = useState(true);
  const [enabledTableNames, setEnabledTableNames] = useState(readEnabledTables);
  const [fieldsByTable, setFieldsByTable] = useState({});
  const [fieldsLoading, setFieldsLoading] = useState(false);
  // Filtered down to CURRENTLY enabled tables, not the whole cache — un-picking a table in the
  // multiselect must make it (and its columns) disappear from the picker right away, even though
  // fieldsByTable itself keeps its columns cached so re-picking it later is instant instead of a
  // fresh fetch (caught live, 2026-09-27: un-picking "Diagnosis" left it sitting in "Selected
  // tables" below since dimensions was built from the whole cache instead of just what's enabled).
  const dimensions = enabledTableNames.flatMap((name) => fieldsByTable[name] ?? []);

  const updateEnabledTables = (names) => {
    setEnabledTableNames(names);
    try {
      localStorage.setItem(ENABLED_TABLES_STORAGE_KEY, JSON.stringify(names));
    } catch {
      // Per-browser convenience only — nothing breaks if this can't be saved (private window,
      // storage disabled, quota); the picker just won't remember the choice next time.
    }
  };

  // Number format and conditional format are pure display settings — computed entirely client-side
  // from cell values already in `result`, no cube round-trip needed. But the grid/chart/export/save
  // below all key off `lastQuery.values` (the frozen snapshot from the last successful query, whose
  // field order/count actually matches `result`'s columns) rather than the live `request.values` —
  // so picking Color scale in a value pill's popover looked like it silently did nothing until you
  // clicked Run query again (caught live, 2026-09-27). Overlaying just `format`/`conditionalFormat`
  // from the live request onto lastQuery's values (matched by field name) fixes that without risking
  // a mismatch when the live request has ADDED/REMOVED a value field entirely — that still correctly
  // needs a fresh Run query, since `result` wouldn't have that field's column yet.
  const displayValues = lastQuery?.values?.map((v) => {
    const live = request.values.find((rv) => rv.field === v.field);
    return live ? { ...v, format: live.format, conditionalFormat: live.conditionalFormat } : v;
  });

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
    Promise.all([getTables(), getMeasures()])
      .then(([tbls, meas]) => {
        setTables(tbls);
        setMeasures(meas);
      })
      .catch((err) => {
        // A rejection here isn't always a plain Error with a real .message (caught live,
        // 2026-09-27: a concurrent-token-refresh race — see dataverseAuth.js's WHY comment —
        // produced one whose falsy .message made {metaError && (...)} render nothing at all,
        // a completely silent failure with no visible error anywhere). Falling back through
        // toString()/JSON.stringify guarantees something is always shown here.
        setMetaError(String(err?.message || err?.toString?.() || JSON.stringify(err) || "Unknown error."));
      })
      .finally(() => setTablesLoading(false));
  }, []);

  // Lazily fetches columns for any ENABLED table not already in the cache — fires once on mount
  // for whatever was already picked in a previous session (from localStorage) and again each time
  // the user adds a table in the multiselect. Never re-fetches a table already cached, and never
  // drops a table's columns from the cache just because it was un-picked (re-picking it later is
  // then instant, and any field from it still placed on the pivot keeps working).
  //
  // enabledTableNames holds DISPLAY names ("Doctor"), matching what the multiselect shows and what
  // FieldPicker's own column-grouping already keys on — but getDimensionsForTable needs the cube's
  // raw identifier ("[Dim Doctor]"), only known once `tables` itself has loaded. Waiting on
  // `tables.length` below avoids caching an empty result for a name whose real table just hasn't
  // arrived yet (that name would otherwise look "already fetched" and never be retried).
  useEffect(() => {
    if (tables.length === 0) return;

    const missing = enabledTableNames.filter((name) => !(name in fieldsByTable));
    if (missing.length === 0) return;

    setFieldsLoading(true);
    Promise.all(
      missing.map((name) => {
        const table = tables.find((t) => t.displayName === name);
        return table ? getDimensionsForTable(table.field).then((fields) => [name, fields]) : Promise.resolve([name, []]);
      }),
    )
      .then((entries) => {
        setFieldsByTable((prev) => {
          const next = { ...prev };
          for (const [name, fields] of entries) next[name] = fields;
          return next;
        });
      })
      .catch((err) => {
        setMetaError(String(err?.message || err?.toString?.() || JSON.stringify(err) || "Unknown error."));
      })
      .finally(() => setFieldsLoading(false));
  }, [enabledTableNames, fieldsByTable, tables]);

  // Drop the table currently on screen the moment the user changes anything that would actually
  // change the query (as opposed to a pure display tweak like Format/Conditional format, which
  // structuralSignature ignores) — otherwise it just sits there looking current while it's really
  // built from Rows/Columns/Filters/etc that no longer match what's selected (explicitly requested,
  // 2026-09-27). Save/export messages are cleared alongside it since they describe that now-gone
  // result too.
  useEffect(() => {
    if (result && structuralSignature(request) !== lastRunSignature) {
      setResult(null);
      setLastQuery(null);
      setSaveMessage(null);
      setExportError(null);
    }
  }, [request, result, lastRunSignature]);

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
      setLastRunSignature(structuralSignature(request));
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

  // Clears the whole builder back to a blank slate — fields, result, every warning/message, saved
  // overrides, and the save-panel inputs. Doesn't touch dimensions/measures (still the same cube,
  // no need to re-fetch metadata) or loading state (a reset mid-query just lets that request finish
  // and get its result silently discarded when it lands, same as any other stale-response case).
  const handleReset = () => {
    setRequest(EMPTY_REQUEST);
    setLastQuery(null);
    setResult(null);
    setError(null);
    setCardinalityWarning(null);
    setCellsWarning(null);
    setConfirmedOverrides({});
    setSnapshotName("");
    setTableStyle("Default");
    setIncludeChart(false);
    setChartType("Bar");
    setReportLayout("Tabular");
    setSaveMessage(null);
    setSaveError(null);
    setExportError(null);
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

  const handleSave = async () => {
    if (!snapshotName.trim() || !result || !lastQuery) {
      return;
    }
    setSaving(true);
    setSaveError(null);
    setSaveMessage(null);
    try {
      const saved = await saveSnapshot(snapshotName.trim(), { ...lastQuery, values: displayValues }, result, tableStyle, includeChart, chartType, reportLayout);
      setSaveMessage(`Saved as version ${saved.versionNumber}.`);
    } catch (err) {
      // SnapshotTooLarge included — a flat, non-overridable ceiling now (removed 2026-09-27 on
      // request), so this just shows like any other save error rather than a confirm-and-retry.
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
        valueFields: displayValues,
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
        <FieldPicker
          value={request}
          onChange={setRequest}
          dimensions={dimensions}
          measures={measures}
          tables={tables}
          tablesLoading={tablesLoading}
          enabledTableNames={enabledTableNames}
          onEnabledTableNamesChange={updateEnabledTables}
          fieldsLoading={fieldsLoading}
          onReset={handleReset}
        />
      </div>

      <button type="button" onClick={() => runQuery()} disabled={loading} className="btn-primary disabled:cursor-not-allowed disabled:opacity-50">
        {loading ? "Running…" : "Run query"}
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
          valueFields={displayValues}
          layout={reportLayout}
          valuesPlacement={lastQuery?.valuesPlacement}
        />
      </div>

      {includeChart &&
        chartUnavailableReason(result, lastQuery) && (
          <p className="mt-4 rounded-xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm text-muted">{chartUnavailableReason(result, lastQuery)}</p>
        )}
      {includeChart && (
        <PivotChart result={result} valueFields={displayValues} chartType={chartType} valuesPlacement={lastQuery?.valuesPlacement} />
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

    </AppShell>
  );
}
