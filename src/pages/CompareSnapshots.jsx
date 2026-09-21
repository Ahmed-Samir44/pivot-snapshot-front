import { useEffect, useState } from "react";
import AppShell from "../components/layout/AppShell";
import { listDefinitions, getVersions, compareVersions } from "../services/snapshotApi";

const nativeSelectClass =
  "rounded-xl border border-slate-200 bg-white/90 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent";

// Flattens every definition's versions into one list of { id, label } options, so either side of
// the comparison can pick any version of any saved pivot — comparison only requires matching
// Rows/Columns/Values (see StructureHasher), not being versions of the SAME definition.
function useAllVersionOptions() {
  const [options, setOptions] = useState([]);

  useEffect(() => {
    listDefinitions().then(async (definitions) => {
      const perDefinition = await Promise.all(
        definitions.map(async (d) => {
          const versions = await getVersions(d.id);
          return versions.map((v) => ({
            value: v.id,
            label: `${d.name} — v${v.versionNumber} (${new Date(v.createdAtUtc).toLocaleDateString()})`,
          }));
        }),
      );
      setOptions(perDefinition.flat());
    });
  }, []);

  return options;
}

export default function CompareSnapshots() {
  const options = useAllVersionOptions();
  const [versionA, setVersionA] = useState("");
  const [versionB, setVersionB] = useState("");
  const [comparison, setComparison] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);

  const runCompare = async () => {
    if (!versionA || !versionB) return;
    setLoading(true);
    setError(null);
    try {
      setComparison(await compareVersions(versionA, versionB));
    } catch (err) {
      setError(err.message);
      setComparison(null);
    } finally {
      setLoading(false);
    }
  };

  return (
    <AppShell>
      <header className="mb-8">
        <h1 className="text-3xl font-bold text-ink">Compare Snapshots</h1>
        <p className="mt-1 text-muted">Pick two saved versions with the same Rows/Columns/Values to see what changed.</p>
      </header>

      <div className="card mb-6 flex flex-wrap items-end gap-3">
        <div>
          <label className="mb-1 block text-xs font-medium text-muted">Version A</label>
          <select className={nativeSelectClass} value={versionA} onChange={(e) => setVersionA(e.target.value)}>
            <option value="">Choose…</option>
            {options.map((o) => (
              <option key={o.value} value={o.value}>
                {o.label}
              </option>
            ))}
          </select>
        </div>
        <div>
          <label className="mb-1 block text-xs font-medium text-muted">Version B</label>
          <select className={nativeSelectClass} value={versionB} onChange={(e) => setVersionB(e.target.value)}>
            <option value="">Choose…</option>
            {options.map((o) => (
              <option key={o.value} value={o.value}>
                {o.label}
              </option>
            ))}
          </select>
        </div>
        <button
          type="button"
          onClick={runCompare}
          disabled={!versionA || !versionB || loading}
          className="btn-primary disabled:cursor-not-allowed disabled:opacity-50"
        >
          {loading ? "Comparing…" : "Compare"}
        </button>
      </div>

      {error && <p className="mb-6 rounded-xl border border-red-300 bg-red-50 px-4 py-3 text-red-800">{error}</p>}

      {comparison && (
        <div className="card overflow-x-auto">
          <table className="w-full border-collapse text-sm">
            <thead>
              <tr>
                <th className="border border-slate-200 bg-gold/10 px-3 py-2 text-left">Row</th>
                <th className="border border-slate-200 bg-gold/10 px-3 py-2 text-left">Column</th>
                <th className="border border-slate-200 bg-gold/10 px-3 py-2 text-right">A</th>
                <th className="border border-slate-200 bg-gold/10 px-3 py-2 text-right">B</th>
                <th className="border border-slate-200 bg-gold/10 px-3 py-2 text-right">Δ</th>
              </tr>
            </thead>
            <tbody>
              {comparison.cells.map((cell, index) => (
                <tr key={index}>
                  <td className="border border-slate-200 px-3 py-2">{cell.rowLabels.join(" / ")}</td>
                  <td className="border border-slate-200 px-3 py-2">{cell.columnLabels.join(" / ")}</td>
                  <td className="border border-slate-200 px-3 py-2 text-right">{cell.valueA ?? "—"}</td>
                  <td className="border border-slate-200 px-3 py-2 text-right">{cell.valueB ?? "—"}</td>
                  <td
                    className={`border border-slate-200 px-3 py-2 text-right font-semibold ${
                      cell.difference > 0 ? "text-green-700" : cell.difference < 0 ? "text-red-700" : ""
                    }`}
                  >
                    {cell.difference ?? "—"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </AppShell>
  );
}
