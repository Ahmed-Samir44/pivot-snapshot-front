import { useState } from "react";
import FieldPicker from "../components/pivot/FieldPicker";
import PivotGrid from "../components/pivot/PivotGrid";
import { queryPivot } from "../services/pivotApi";

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

// Filter members are edited as one comma-separated text field in the UI (simplest input for an
// MVP member picker) but the API wants an array — this is the only place that split happens.
function parseMembers(membersText) {
  return membersText
    .split(",")
    .map((m) => m.trim())
    .filter(Boolean);
}

export default function PivotBuilder() {
  const [request, setRequest] = useState(EMPTY_REQUEST);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);

  const runQuery = async () => {
    setLoading(true);
    setError(null);
    try {
      const cleaned = {
        rows: request.rows.filter(Boolean),
        columns: request.columns.filter(Boolean),
        values: request.values.filter((v) => v.field),
        filters: request.filters
          .filter((f) => f.field)
          .map((f) => ({ field: f.field, includedMembers: parseMembers(f.membersText) }))
          .filter((f) => f.includedMembers.length > 0),
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
    <main className="pivot-builder">
      <h1>Pivot Snapshot Builder</h1>
      <FieldPicker value={request} onChange={setRequest} />
      <button type="button" onClick={runQuery} disabled={loading}>
        {loading ? "جاري التحميل..." : "عرض الجدول"}
      </button>
      {error && <p className="pivot-error">{error}</p>}
      <PivotGrid result={result} />
    </main>
  );
}
