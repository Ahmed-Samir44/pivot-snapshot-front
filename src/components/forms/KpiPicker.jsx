import { useEffect, useState } from "react";
import { Loader2, Search } from "lucide-react";
import { getKpis } from "../../services/kpiApi";

// A searchable KPI list, reused by both "Link to KPI…" (Snapshot History) and "Choose KPI…"
// (Builder) — fetches stf_strategykpis itself (small, read-only) rather than asking each caller to
// pass the list down, since both call sites would otherwise duplicate the same fetch/loading state.
export default function KpiPicker({ onSelect }) {
  const [kpis, setKpis] = useState(null);
  const [error, setError] = useState(null);
  const [search, setSearch] = useState("");

  useEffect(() => {
    let cancelled = false;
    getKpis()
      .then((data) => !cancelled && setKpis(data))
      .catch((err) => !cancelled && setError(err.message));
    return () => {
      cancelled = true;
    };
  }, []);

  const query = search.trim().toLowerCase();
  const visible = (kpis ?? []).filter((k) => k.name.toLowerCase().includes(query));

  return (
    <div className="flex w-72 flex-col gap-2">
      <div className="relative">
        <Search className="pointer-events-none absolute left-2.5 top-1/2 h-3.5 w-3.5 -translate-y-1/2 text-muted" />
        <input
          type="text"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          placeholder="Search KPIs…"
          className="w-full rounded-md border border-slate-200 bg-white py-1.5 pl-8 pr-3 text-xs focus:outline-none focus:ring-2 focus:ring-primary focus:border-transparent"
        />
      </div>
      <div className="max-h-64 overflow-y-auto rounded-lg border border-slate-200">
        {kpis === null && !error && (
          <p className="flex items-center gap-1.5 px-2 py-2 text-xs text-muted">
            <Loader2 className="h-3.5 w-3.5 animate-spin" /> Loading KPIs…
          </p>
        )}
        {error && <p className="px-2 py-2 text-xs text-red-700">{error}</p>}
        {kpis !== null && visible.length === 0 && <p className="px-2 py-2 text-xs text-muted">No matching KPIs.</p>}
        {visible.map((kpi) => (
          <button
            key={kpi.id}
            type="button"
            onClick={() => onSelect(kpi)}
            className="block w-full truncate px-2 py-1.5 text-left text-xs text-ink hover:bg-slate-100"
          >
            {kpi.name}
          </button>
        ))}
      </div>
    </div>
  );
}
