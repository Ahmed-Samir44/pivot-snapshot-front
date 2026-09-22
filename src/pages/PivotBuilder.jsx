import { useEffect, useState } from "react";
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
  filters: [],
  sort: null,
  showGrandTotals: false,
  showSubtotals: false,
  calculatedFields: [],
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
  const [saving, setSaving] = useState(false);
  const [saveMessage, setSaveMessage] = useState(null);
  const [saveError, setSaveError] = useState(null);

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
        filters: request.filters.filter((f) => f.field && isFilterComplete(f)),
        sort: request.sort,
        showGrandTotals: request.showGrandTotals,
        showSubtotals: request.showSubtotals,
        calculatedFields: request.calculatedFields.filter((f) => f.name && f.leftField && f.rightField),
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
      const saved = await saveSnapshot(snapshotName.trim(), lastQuery, result, tableStyle, includeChart, chartType);
      setSaveMessage(`Saved as version ${saved.versionNumber}.`);
    } catch (err) {
      setSaveError(err.message);
    } finally {
      setSaving(false);
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
                السيرفر ده مشترك مع مشاريع تانية وسبق وحصله انهيار بسبب استعلام مشابه — تشغيل الاستعلام ده من غير فلتر ممكن يبطّئ أو يعطّل السيرفر للمستخدمين التانيين. اضغط بس لو متأكد.
              </p>
              <button type="button" onClick={runQueryWithOverride} disabled={loading} className="btn-secondary mt-2 disabled:cursor-not-allowed disabled:opacity-50">
                فهمت المخاطرة، شغّل على أي حال (لحد {cardinalityWarning.absoluteMaxMembersPerField.toLocaleString()} عضو)
              </button>
            </>
          ) : (
            <p className="mt-1 text-sm">
              عدد الأعضاء ({cardinalityWarning.memberCount.toLocaleString()}) أكبر من الحد الأقصى المطلق المسموح به على السيرفر ({cardinalityWarning.absoluteMaxMembersPerField.toLocaleString()}) — لازم تضيف فلتر على الحقل ده، مفيش تجاوز ممكن هنا.
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
        />
      </div>

      {includeChart && <PivotChart result={result} valueFields={lastQuery?.values} chartType={chartType} />}

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
          {saveMessage && <span className="text-sm font-medium text-green-700">{saveMessage}</span>}
          {saveError && <span className="text-sm font-medium text-red-700">{saveError}</span>}
        </div>
      )}
    </AppShell>
  );
}
