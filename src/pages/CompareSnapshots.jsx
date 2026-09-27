import { useEffect, useState } from "react";
import { Plus, X } from "lucide-react";
import AppShell from "../components/layout/AppShell";
import { listDefinitions, getVersions, compareVersions } from "../services/snapshotApi";

const nativeSelectClass =
  "rounded-xl border border-slate-200 bg-white/90 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-gold focus:border-transparent";

// Flattens every definition's versions into one list of { id, label } options, so any slot can
// pick any version of any saved pivot — comparison only requires matching Rows/Columns/Values (see
// StructureHasher), not being versions of the SAME definition.
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

// 2026-09-24: generalized from exactly 2 fixed slots (Version A / Version B) to N (2+) — versions
// is now an array of selected ids, with Add/Remove controls instead of two hardcoded selects.
export default function CompareSnapshots() {
  const options = useAllVersionOptions();
  const [versions, setVersions] = useState(["", ""]);
  const [comparison, setComparison] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);

  const allFilled = versions.every(Boolean);

  const setVersionAt = (index, id) => {
    setVersions(versions.map((v, i) => (i === index ? id : v)));
  };

  const addSlot = () => setVersions([...versions, ""]);

  const removeSlot = (index) => setVersions(versions.filter((_, i) => i !== index));

  const runCompare = async () => {
    if (!allFilled) return;
    setLoading(true);
    setError(null);
    try {
      setComparison(await compareVersions(versions));
    } catch (err) {
      setError(err.message);
      setComparison(null);
    } finally {
      setLoading(false);
    }
  };

  const labelFor = (id) => options.find((o) => o.value === id)?.label ?? id;

  return (
    <AppShell>
      <header className="mb-8">
        <h1 className="text-3xl font-bold text-ink">Compare Snapshots</h1>
        <p className="mt-1 text-muted">Pick 2 or more saved versions with the same Rows/Columns/Values to see what changed.</p>
      </header>

      <div className="card mb-6 flex flex-wrap items-end gap-3">
        {versions.map((v, index) => (
          <div key={index}>
            <label className="mb-1 flex items-center gap-1.5 text-xs font-medium text-muted">
              Version {index + 1}
              {versions.length > 2 && (
                <button type="button" onClick={() => removeSlot(index)} className="text-muted hover:text-ink" aria-label={`Remove version ${index + 1}`}>
                  <X className="h-3.5 w-3.5" />
                </button>
              )}
            </label>
            <select className={nativeSelectClass} value={v} onChange={(e) => setVersionAt(index, e.target.value)}>
              <option value="">Choose…</option>
              {options.map((o) => (
                <option key={o.value} value={o.value}>
                  {o.label}
                </option>
              ))}
            </select>
          </div>
        ))}
        <button type="button" onClick={addSlot} className="btn-secondary inline-flex items-center gap-1.5 px-3 py-2 text-sm">
          <Plus className="h-4 w-4" /> Add version
        </button>
        <button
          type="button"
          onClick={runCompare}
          disabled={!allFilled || loading}
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
                {versions.map((id, index) => (
                  <th key={id} className="border border-slate-200 bg-gold/10 px-3 py-2 text-right" title={labelFor(id)}>
                    V{index + 1}
                  </th>
                ))}
                {versions.slice(1).map((_, index) => (
                  <th key={`delta-${index}`} className="border border-slate-200 bg-gold/10 px-3 py-2 text-right">
                    Δ (V{index + 1}→V{index + 2})
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {comparison.cells.map((cell, index) => (
                <tr key={index}>
                  <td className="border border-slate-200 px-3 py-2">{cell.rowLabels.join(" / ")}</td>
                  <td className="border border-slate-200 px-3 py-2">{cell.columnLabels.join(" / ")}</td>
                  {cell.values.map((value, i) => (
                    <td key={i} className="border border-slate-200 px-3 py-2 text-right">
                      {value ?? "—"}
                    </td>
                  ))}
                  {cell.deltas.map((delta, i) => (
                    <td
                      key={i}
                      className={`border border-slate-200 px-3 py-2 text-right font-semibold ${
                        delta > 0 ? "text-green-700" : delta < 0 ? "text-red-700" : ""
                      }`}
                    >
                      {delta ?? "—"}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </AppShell>
  );
}
