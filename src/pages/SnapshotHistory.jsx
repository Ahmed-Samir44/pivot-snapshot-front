import { useEffect, useState } from "react";
import AppShell from "../components/layout/AppShell";
import { listDefinitions, getVersions, getVersion } from "../services/snapshotApi";

export default function SnapshotHistory() {
  const [definitions, setDefinitions] = useState([]);
  const [selectedDefinitionId, setSelectedDefinitionId] = useState(null);
  const [versions, setVersions] = useState([]);
  const [selectedVersionHtml, setSelectedVersionHtml] = useState(null);
  const [error, setError] = useState(null);

  useEffect(() => {
    listDefinitions()
      .then(setDefinitions)
      .catch((err) => setError(err.message));
  }, []);

  const selectDefinition = async (id) => {
    setSelectedDefinitionId(id);
    setSelectedVersionHtml(null);
    try {
      setVersions(await getVersions(id));
    } catch (err) {
      setError(err.message);
    }
  };

  const selectVersion = async (versionId) => {
    try {
      const version = await getVersion(versionId);
      setSelectedVersionHtml(version.htmlContent);
    } catch (err) {
      setError(err.message);
    }
  };

  return (
    <AppShell>
      <header className="mb-8">
        <h1 className="text-3xl font-bold text-ink">Snapshot History</h1>
        <p className="mt-1 text-muted">Every saved pivot, and every version saved under it.</p>
      </header>

      {error && <p className="mb-6 rounded-xl border border-red-300 bg-red-50 px-4 py-3 text-red-800">{error}</p>}

      <div className="grid grid-cols-1 gap-6 lg:grid-cols-3">
        <div className="card lg:col-span-1">
          <h3 className="mb-4 text-sm font-bold uppercase tracking-wide text-muted">Saved Pivots</h3>
          <ul className="space-y-2">
            {definitions.map((d) => (
              <li key={d.id}>
                <button
                  type="button"
                  onClick={() => selectDefinition(d.id)}
                  className={`w-full rounded-xl px-4 py-3 text-left transition-colors ${
                    selectedDefinitionId === d.id ? "bg-gold/20 font-semibold" : "hover:bg-slate-100"
                  }`}
                >
                  {d.name}
                  <span className="ml-2 text-xs text-muted">
                    {d.versionCount} version{d.versionCount === 1 ? "" : "s"}
                  </span>
                </button>
              </li>
            ))}
            {definitions.length === 0 && <li className="text-sm text-muted">No snapshots saved yet.</li>}
          </ul>
        </div>

        <div className="card lg:col-span-1">
          <h3 className="mb-4 text-sm font-bold uppercase tracking-wide text-muted">Versions</h3>
          <ul className="space-y-2">
            {versions.map((v) => (
              <li key={v.id}>
                <button
                  type="button"
                  onClick={() => selectVersion(v.id)}
                  className="w-full rounded-xl px-4 py-3 text-left hover:bg-slate-100"
                >
                  Version {v.versionNumber}
                  <span className="ml-2 text-xs text-muted">{new Date(v.createdAtUtc).toLocaleString()}</span>
                </button>
              </li>
            ))}
            {selectedDefinitionId && versions.length === 0 && <li className="text-sm text-muted">No versions.</li>}
          </ul>
        </div>

        <div className="card lg:col-span-1 overflow-x-auto">
          <h3 className="mb-4 text-sm font-bold uppercase tracking-wide text-muted">Preview</h3>
          {selectedVersionHtml ? (
            // Trusted content: this HTML is rendered server-side by our own HtmlSnapshotRenderer,
            // not arbitrary third-party or user-supplied markup.
            <div dangerouslySetInnerHTML={{ __html: selectedVersionHtml }} />
          ) : (
            <p className="text-sm text-muted">Pick a version to preview it.</p>
          )}
        </div>
      </div>
    </AppShell>
  );
}
