import { useEffect, useState } from "react";
import AppShell from "../components/layout/AppShell";
import FieldPicker from "../components/pivot/FieldPicker";
import PivotGrid from "../components/pivot/PivotGrid";
import { queryPivot } from "../services/pivotApi";
import { getDimensions, getMeasures } from "../services/cubeMetaApi";

// This screen renders a live, interactive pivot (React state + re-fetch on every change).
// It is NOT the saved snapshot: "Save Snapshot" (not built yet — see DECISIONS.md Phase C) will
// freeze the current result into static HTML/CSS for Power Apps, which cannot run JavaScript.
const EMPTY_REQUEST = {
  rows: [],
  columns: [],
  values: [],
  filters: [],
  sort: null,
  showGrandTotals: false,
};

export default function PivotBuilder() {
  const [request, setRequest] = useState(EMPTY_REQUEST);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [dimensions, setDimensions] = useState([]);
  const [measures, setMeasures] = useState([]);
  const [metaError, setMetaError] = useState(null);

  useEffect(() => {
    Promise.all([getDimensions(), getMeasures()])
      .then(([dims, meas]) => {
        setDimensions(dims);
        setMeasures(meas);
      })
      .catch((err) => setMetaError(err.message));
  }, []);

  const runQuery = async () => {
    setLoading(true);
    setError(null);
    try {
      const cleaned = {
        rows: request.rows.filter(Boolean),
        columns: request.columns.filter(Boolean),
        values: request.values.filter((v) => v.field),
        filters: request.filters.filter((f) => f.field && f.includedMembers.length > 0),
        sort: request.sort,
        showGrandTotals: request.showGrandTotals,
      };
      const data = await queryPivot(cleaned);
      setResult(data);
    } catch (err) {
      setError(err.message);
      setResult(null);
    } finally {
      setLoading(false);
    }
  };

  return (
    <AppShell>
      <header className="mb-8">
        <h1 className="text-3xl font-bold text-ink">Pivot Snapshot Builder</h1>
        <p className="mt-1 text-muted">Build a pivot, then (soon) save it as a snapshot.</p>
      </header>

      {metaError && (
        <p className="mb-6 rounded-xl border border-amber-300 bg-amber-50 px-4 py-3 text-amber-800">
          Couldn&apos;t load the field list from the API: {metaError}
        </p>
      )}

      <div className="mb-6">
        <FieldPicker value={request} onChange={setRequest} dimensions={dimensions} measures={measures} />
      </div>

      <button type="button" onClick={runQuery} disabled={loading} className="btn-primary disabled:cursor-not-allowed disabled:opacity-50">
        {loading ? "Loading…" : "Run query"}
      </button>

      {error && (
        <p className="mt-4 rounded-xl border border-red-300 bg-red-50 px-4 py-3 text-red-800">{error}</p>
      )}

      <div className="mt-6">
        <PivotGrid result={result} />
      </div>
    </AppShell>
  );
}
