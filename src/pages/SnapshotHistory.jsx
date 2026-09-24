import { useEffect, useState } from "react";
import { ChevronRight, CirclePause, CirclePlay, Eye, FileSpreadsheet, Trash2, X } from "lucide-react";
import AppShell from "../components/layout/AppShell";
import Modal from "../components/forms/Modal";
import ConfirmModal from "../components/forms/ConfirmModal";
import PivotGrid from "../components/pivot/PivotGrid";
import {
  listDefinitions,
  getDeactivatedDefinitions,
  deactivateDefinition,
  reactivateDefinition,
  deleteDefinition,
  getVersions,
  getDeactivatedVersions,
  getVersion,
  deactivateVersion,
  reactivateVersion,
  reactivateAllVersions,
  deleteVersion,
} from "../services/snapshotApi";
import { formatDimensionName } from "../utils/cubeMeta";

// A plain Set-backed multi-select for one list — its own little state machine (toggle one, toggle
// all, clear) reused across all four lists (Active/Deactivated × Pivots/Versions) instead of
// writing the same three functions four times.
function useSelection() {
  const [ids, setIds] = useState(new Set());

  const toggle = (id) =>
    setIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });

  const clear = () => setIds(new Set());

  return { ids, toggle, clear };
}

// One row — used for pivots and versions alike, both Active and Deactivated lists — a bordered
// card per row (not just a hover tint) so rows read as distinct items instead of a dense block of
// text and icons running into each other (caught live 2026-09-23: "الشكل كله على بعضه").
function ListRow({ label, sublabel, onClick, trailingIcon: TrailingIcon, actions, highlighted = false, checked, onToggleCheck }) {
  return (
    <li className="flex items-center gap-1 rounded-xl border border-slate-100 py-1 pl-2 pr-2">
      <input
        type="checkbox"
        checked={checked}
        onChange={onToggleCheck}
        className="h-4 w-4 shrink-0 accent-gold"
        aria-label={`Select ${label}`}
      />
      <button
        type="button"
        onClick={onClick}
        disabled={!onClick}
        className={`flex flex-1 items-center justify-between gap-2 rounded-lg px-3 py-2.5 text-left disabled:cursor-default disabled:hover:bg-transparent ${
          highlighted ? "bg-gold/20 font-semibold" : "hover:bg-slate-50"
        }`}
      >
        <span>
          <span className="font-medium text-ink">{label}</span>
          {sublabel && <span className="ml-2 text-xs text-muted">{sublabel}</span>}
        </span>
        {TrailingIcon && <TrailingIcon className="h-4 w-4 shrink-0 text-muted" aria-hidden="true" />}
      </button>
      <div className="flex shrink-0 items-center gap-1">{actions}</div>
    </li>
  );
}

function SectionLabel({ color, children, action }) {
  return (
    <div className="mb-2 flex items-center justify-between gap-2">
      <p className="flex items-center gap-1.5 text-xs font-bold uppercase tracking-wide text-muted">
        <span className={`h-2 w-2 rounded-full ${color}`} aria-hidden="true" />
        {children}
      </p>
      {action}
    </div>
  );
}

// The toolbar that appears above a list once at least one row is checked — "N selected" plus
// whichever bulk actions apply to THIS list (an active list only offers Deactivate; a deactivated
// list offers Reactivate and Delete), so the same component covers all four lists.
function BulkActionBar({ selection, onDeactivate, onReactivate, onDelete }) {
  if (selection.ids.size === 0) return null;
  return (
    <div className="mb-2 flex items-center gap-3 rounded-lg bg-slate-50 px-3 py-2">
      <span className="text-xs font-semibold text-ink">{selection.ids.size} selected</span>
      {onDeactivate && (
        <button type="button" onClick={onDeactivate} className="text-xs font-semibold text-orange-600 hover:underline">
          Deactivate
        </button>
      )}
      {onReactivate && (
        <button type="button" onClick={onReactivate} className="text-xs font-semibold text-green-700 hover:underline">
          Reactivate
        </button>
      )}
      {onDelete && (
        <button type="button" onClick={onDelete} className="text-xs font-semibold text-red-700 hover:underline">
          Delete permanently
        </button>
      )}
      <button type="button" onClick={selection.clear} className="ml-auto text-muted hover:text-ink" aria-label="Clear selection">
        <X className="h-3.5 w-3.5" />
      </button>
    </div>
  );
}

export default function SnapshotHistory() {
  const [definitions, setDefinitions] = useState([]);
  const [deactivatedDefinitions, setDeactivatedDefinitions] = useState([]);
  const [selectedDefinitionId, setSelectedDefinitionId] = useState(null);
  const [versions, setVersions] = useState([]);
  // Deactivated versions for the currently-selected pivot — a "recycle bin" list, fetched
  // alongside the normal active list whenever a pivot is selected. Never mixed into `versions`.
  const [deactivatedVersions, setDeactivatedVersions] = useState([]);
  const [error, setError] = useState(null);

  // The interactive popup (search/group-collapse via the live PivotGrid) is the ONLY preview now
  // — the earlier static third-column preview was dropped (2026-09-23) once this replaced it,
  // since showing the same version two different ways was just redundant.
  const [previewVersion, setPreviewVersion] = useState(null);
  const [previewError, setPreviewError] = useState(null);
  const [exporting, setExporting] = useState(false);
  const [exportError, setExportError] = useState(null);

  // One checkbox selection per list — a checked row in "Active Versions" has nothing to do with
  // one checked in "Deactivated Versions", so these never share state (caught live 2026-09-23:
  // "مش كلهم" — the user wants to pick an arbitrary subset, not just "all" or "one").
  const activeDefSelection = useSelection();
  const deactivatedDefSelection = useSelection();
  const activeVersionSelection = useSelection();
  const deactivatedVersionSelection = useSelection();

  // Replaces window.confirm() (some environments block/discourage it, flagged live 2026-09-23) —
  // { action, ids: string[], label } while a confirmation is pending, null otherwise. Always an
  // array of ids, even for a single-row action, so one confirm/handleConfirm path covers both.
  const [pendingConfirm, setPendingConfirm] = useState(null);

  const refreshDefinitionLists = async () => {
    const [active, deactivated] = await Promise.all([listDefinitions(), getDeactivatedDefinitions()]);
    setDefinitions(active);
    setDeactivatedDefinitions(deactivated);
  };

  useEffect(() => {
    listDefinitions()
      .then(setDefinitions)
      .catch((err) => setError(err.message));
    getDeactivatedDefinitions()
      .then(setDeactivatedDefinitions)
      .catch((err) => setError(err.message));
  }, []);

  const refreshVersionLists = async (definitionId) => {
    const [active, deactivated] = await Promise.all([getVersions(definitionId), getDeactivatedVersions(definitionId)]);
    setVersions(active);
    setDeactivatedVersions(deactivated);
  };

  const selectDefinition = async (id) => {
    setSelectedDefinitionId(id);
    // A version id checked under the PREVIOUS pivot means nothing once its list is replaced —
    // switching pivots always starts the version selections fresh.
    activeVersionSelection.clear();
    deactivatedVersionSelection.clear();
    try {
      await refreshVersionLists(id);
    } catch (err) {
      setError(err.message);
    }
  };

  const deselectIfCurrent = (definitionIds) => {
    if (definitionIds.includes(selectedDefinitionId)) {
      setSelectedDefinitionId(null);
      setVersions([]);
      setDeactivatedVersions([]);
    }
  };

  const openPreview = async (versionId, versionNumber) => {
    setPreviewError(null);
    setExportError(null);
    try {
      const version = await getVersion(versionId);
      setPreviewVersion({ ...version, versionNumber });
    } catch (err) {
      setPreviewError(err.message);
    }
  };

  // fn is one of deactivateVersion/reactivateVersion/deleteVersion — versionIds is always an
  // array (length 1 for a single-row action), run in parallel since these are plain Dataverse
  // column updates, not cube queries (no load-on-the-shared-server concern here).
  const runVersionAction = async (fn, versionIds) => {
    setPreviewError(null);
    try {
      await Promise.all(versionIds.map((id) => fn(id)));
      if (versionIds.includes(previewVersion?.id)) {
        setPreviewVersion(null);
      }
      await Promise.all([refreshVersionLists(selectedDefinitionId), refreshDefinitionLists()]);
      activeVersionSelection.clear();
      deactivatedVersionSelection.clear();
    } catch (err) {
      setPreviewError(err.message);
    }
  };

  const handleReactivateAll = async () => {
    setPreviewError(null);
    try {
      await reactivateAllVersions(selectedDefinitionId);
      await Promise.all([refreshVersionLists(selectedDefinitionId), refreshDefinitionLists()]);
    } catch (err) {
      setPreviewError(err.message);
    }
  };

  const runDefinitionAction = async (fn, definitionIds) => {
    setError(null);
    try {
      await Promise.all(definitionIds.map((id) => fn(id)));
      deselectIfCurrent(definitionIds);
      await refreshDefinitionLists();
      activeDefSelection.clear();
      deactivatedDefSelection.clear();
    } catch (err) {
      setError(err.message);
    }
  };

  const handleConfirm = async () => {
    const { action, ids } = pendingConfirm;
    setPendingConfirm(null);
    if (action === "deactivate-version") await runVersionAction(deactivateVersion, ids);
    if (action === "delete-version") await runVersionAction(deleteVersion, ids);
    if (action === "deactivate-definition") await runDefinitionAction(deactivateDefinition, ids);
    if (action === "delete-definition") await runDefinitionAction(deleteDefinition, ids);
  };

  const countLabel = (count, singular) => (count === 1 ? singular : `${count} ${singular}s`);

  const handleExport = async () => {
    if (!previewVersion) return;
    setExporting(true);
    setExportError(null);
    try {
      // Dynamic import — see PivotBuilder.jsx's handleExport for why (ExcelJS is ~1MB, only
      // worth loading for someone who actually clicks Export).
      const { exportPivotToExcel } = await import("../utils/exportExcel");
      const pivotName = definitions.find((d) => d.id === selectedDefinitionId)?.name ?? "Pivot";
      await exportPivotToExcel({
        result: previewVersion.result,
        rowFieldLabels: previewVersion.rows.map(formatDimensionName),
        columnFieldLabels: previewVersion.columns.map(formatDimensionName),
        valueFields: previewVersion.values,
        valuesPlacement: previewVersion.valuesPlacement,
        filename: `${pivotName}_v${previewVersion.versionNumber}`,
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
        <h1 className="text-3xl font-bold text-ink">Snapshot History</h1>
        <p className="mt-1 text-muted">Every saved pivot, and every version saved under it.</p>
      </header>

      {error && <p className="mb-6 rounded-xl border border-red-300 bg-red-50 px-4 py-3 text-red-800">{error}</p>}

      <div className="grid grid-cols-1 gap-6 lg:grid-cols-2">
        <div className="card">
          <h3 className="mb-4 text-sm font-bold uppercase tracking-wide text-muted">Saved Pivots</h3>

          <SectionLabel color="bg-green-500">Active</SectionLabel>
          <BulkActionBar
            selection={activeDefSelection}
            onDeactivate={() => {
              const ids = [...activeDefSelection.ids];
              const versionCount = definitions.filter((d) => ids.includes(d.id)).reduce((sum, d) => sum + d.versionCount, 0);
              setPendingConfirm({ action: "deactivate-definition", ids, label: countLabel(ids.length, "pivot"), versionCount });
            }}
          />
          <ul className="space-y-2">
            {definitions.map((d) => (
              <ListRow
                key={d.id}
                label={d.name}
                sublabel={`${d.versionCount} version${d.versionCount === 1 ? "" : "s"}`}
                onClick={() => selectDefinition(d.id)}
                trailingIcon={ChevronRight}
                highlighted={selectedDefinitionId === d.id}
                checked={activeDefSelection.ids.has(d.id)}
                onToggleCheck={() => activeDefSelection.toggle(d.id)}
                actions={
                  <button
                    type="button"
                    onClick={() =>
                      setPendingConfirm({
                        action: "deactivate-definition",
                        ids: [d.id],
                        label: d.name,
                        versionCount: d.versionCount,
                      })
                    }
                    className="rounded-full p-2 text-orange-500 hover:bg-orange-50 hover:text-orange-700"
                    title="Deactivate this pivot"
                    aria-label={`Deactivate pivot ${d.name}`}
                  >
                    <CirclePause className="h-4 w-4" />
                  </button>
                }
              />
            ))}
            {definitions.length === 0 && <li className="text-sm text-muted">No snapshots saved yet.</li>}
          </ul>

          {deactivatedDefinitions.length > 0 && (
            <div className="mt-6 border-t border-slate-100 pt-5">
              <SectionLabel color="bg-orange-500">Deactivated</SectionLabel>
              <BulkActionBar
                selection={deactivatedDefSelection}
                onReactivate={() => runDefinitionAction(reactivateDefinition, [...deactivatedDefSelection.ids])}
                onDelete={() => {
                  const ids = [...deactivatedDefSelection.ids];
                  setPendingConfirm({ action: "delete-definition", ids, label: countLabel(ids.length, "pivot") });
                }}
              />
              <ul className="space-y-2">
                {deactivatedDefinitions.map((d) => (
                  <ListRow
                    key={d.id}
                    label={d.name}
                    sublabel={`${d.versionCount} version${d.versionCount === 1 ? "" : "s"}`}
                    checked={deactivatedDefSelection.ids.has(d.id)}
                    onToggleCheck={() => deactivatedDefSelection.toggle(d.id)}
                    actions={
                      <>
                        <button
                          type="button"
                          onClick={() => runDefinitionAction(reactivateDefinition, [d.id])}
                          className="rounded-full p-2 text-green-600 hover:bg-green-50 hover:text-green-700"
                          title="Reactivate this pivot"
                          aria-label={`Reactivate pivot ${d.name}`}
                        >
                          <CirclePlay className="h-4 w-4" />
                        </button>
                        <button
                          type="button"
                          onClick={() => setPendingConfirm({ action: "delete-definition", ids: [d.id], label: d.name })}
                          className="rounded-full p-2 text-muted hover:bg-red-50 hover:text-red-700"
                          title="Delete this pivot permanently"
                          aria-label={`Delete pivot ${d.name} permanently`}
                        >
                          <Trash2 className="h-4 w-4" />
                        </button>
                      </>
                    }
                  />
                ))}
              </ul>
            </div>
          )}
        </div>

        <div className="card">
          <h3 className="mb-4 text-sm font-bold uppercase tracking-wide text-muted">Versions</h3>
          {previewError && <p className="mb-3 rounded-xl border border-red-300 bg-red-50 px-3 py-2 text-sm text-red-800">{previewError}</p>}

          <SectionLabel color="bg-green-500">Active</SectionLabel>
          <BulkActionBar
            selection={activeVersionSelection}
            onDeactivate={() => {
              const ids = [...activeVersionSelection.ids];
              setPendingConfirm({ action: "deactivate-version", ids, label: countLabel(ids.length, "version") });
            }}
          />
          <ul className="space-y-2">
            {versions.map((v) => (
              <ListRow
                key={v.id}
                label={`Version ${v.versionNumber}`}
                sublabel={new Date(v.createdAtUtc).toLocaleString()}
                onClick={() => openPreview(v.id, v.versionNumber)}
                trailingIcon={Eye}
                checked={activeVersionSelection.ids.has(v.id)}
                onToggleCheck={() => activeVersionSelection.toggle(v.id)}
                actions={
                  <button
                    type="button"
                    onClick={() => setPendingConfirm({ action: "deactivate-version", ids: [v.id], label: `Version ${v.versionNumber}` })}
                    className="rounded-full p-2 text-orange-500 hover:bg-orange-50 hover:text-orange-700"
                    title="Deactivate this version"
                    aria-label={`Deactivate version ${v.versionNumber}`}
                  >
                    <CirclePause className="h-4 w-4" />
                  </button>
                }
              />
            ))}
            {!selectedDefinitionId && <li className="text-sm text-muted">Pick a saved pivot to see its versions.</li>}
            {selectedDefinitionId && versions.length === 0 && <li className="text-sm text-muted">No active versions.</li>}
          </ul>

          {selectedDefinitionId && deactivatedVersions.length > 0 && (
            <div className="mt-6 border-t border-slate-100 pt-5">
              <SectionLabel
                color="bg-orange-500"
                action={
                  deactivatedVersions.length > 1 && (
                    <button type="button" onClick={handleReactivateAll} className="text-xs font-semibold text-green-700 hover:underline">
                      Reactivate all ({deactivatedVersions.length})
                    </button>
                  )
                }
              >
                Deactivated
              </SectionLabel>
              <BulkActionBar
                selection={deactivatedVersionSelection}
                onReactivate={() => runVersionAction(reactivateVersion, [...deactivatedVersionSelection.ids])}
                onDelete={() => {
                  const ids = [...deactivatedVersionSelection.ids];
                  setPendingConfirm({ action: "delete-version", ids, label: countLabel(ids.length, "version") });
                }}
              />
              <ul className="space-y-2">
                {deactivatedVersions.map((v) => (
                  <ListRow
                    key={v.id}
                    label={`Version ${v.versionNumber}`}
                    sublabel={new Date(v.createdAtUtc).toLocaleString()}
                    onClick={() => openPreview(v.id, v.versionNumber)}
                    trailingIcon={Eye}
                    checked={deactivatedVersionSelection.ids.has(v.id)}
                    onToggleCheck={() => deactivatedVersionSelection.toggle(v.id)}
                    actions={
                      <>
                        <button
                          type="button"
                          onClick={() => runVersionAction(reactivateVersion, [v.id])}
                          className="rounded-full p-2 text-green-600 hover:bg-green-50 hover:text-green-700"
                          title="Reactivate this version"
                          aria-label={`Reactivate version ${v.versionNumber}`}
                        >
                          <CirclePlay className="h-4 w-4" />
                        </button>
                        <button
                          type="button"
                          onClick={() => setPendingConfirm({ action: "delete-version", ids: [v.id], label: `Version ${v.versionNumber}` })}
                          className="rounded-full p-2 text-muted hover:bg-red-50 hover:text-red-700"
                          title="Delete this version permanently"
                          aria-label={`Delete version ${v.versionNumber} permanently`}
                        >
                          <Trash2 className="h-4 w-4" />
                        </button>
                      </>
                    }
                  />
                ))}
              </ul>
            </div>
          )}
        </div>
      </div>

      {pendingConfirm?.action === "deactivate-version" && (
        <ConfirmModal
          title="Deactivate version"
          message={`Deactivate ${pendingConfirm.label}? It'll move to the Deactivated list (recoverable in Dataverse directly, or from there).`}
          confirmLabel="Deactivate"
          onConfirm={handleConfirm}
          onCancel={() => setPendingConfirm(null)}
        />
      )}
      {pendingConfirm?.action === "delete-version" && (
        <ConfirmModal
          title="Delete version permanently"
          message={`Permanently delete ${pendingConfirm.label}? This cannot be undone from anywhere, including Dataverse.`}
          confirmLabel="Delete permanently"
          danger
          onConfirm={handleConfirm}
          onCancel={() => setPendingConfirm(null)}
        />
      )}
      {pendingConfirm?.action === "deactivate-definition" && (
        <ConfirmModal
          title="Deactivate pivot"
          message={
            pendingConfirm.versionCount > 0
              ? `Deactivate ${pendingConfirm.label}? This will ALSO deactivate all ${pendingConfirm.versionCount} version(s) under ${
                  pendingConfirm.ids.length === 1 ? "it" : "them"
                }. Everything moves to the Deactivated list (recoverable from there, or in Dataverse directly).`
              : `Deactivate ${pendingConfirm.label}? ${
                  pendingConfirm.ids.length === 1 ? "It'll" : "They'll"
                } move to the Deactivated list (recoverable from there, or in Dataverse directly).`
          }
          confirmLabel="Deactivate"
          onConfirm={handleConfirm}
          onCancel={() => setPendingConfirm(null)}
        />
      )}
      {pendingConfirm?.action === "delete-definition" && (
        <ConfirmModal
          title="Delete pivot permanently"
          message={`Permanently delete ${pendingConfirm.label} AND every version saved under ${
            pendingConfirm.ids.length === 1 ? "it" : "them"
          }? This cannot be undone from anywhere, including Dataverse.`}
          confirmLabel="Delete permanently"
          danger
          onConfirm={handleConfirm}
          onCancel={() => setPendingConfirm(null)}
        />
      )}

      {previewVersion && (
        <Modal title={`Version ${previewVersion.versionNumber} — interactive preview`} onClose={() => setPreviewVersion(null)}>
          <div className="mb-3 flex flex-wrap items-center gap-3">
            <button
              type="button"
              onClick={handleExport}
              disabled={exporting}
              className="btn-secondary inline-flex items-center gap-1.5 disabled:cursor-not-allowed disabled:opacity-50"
            >
              <FileSpreadsheet className="h-4 w-4" style={{ color: "#217346" }} />
              {exporting ? "Exporting…" : "Export to Excel"}
            </button>
            {exportError && <span className="text-sm font-medium text-red-700">{exportError}</span>}
          </div>
          <PivotGrid
            result={previewVersion.result}
            rowFieldLabels={previewVersion.rows.map(formatDimensionName)}
            columnFieldLabels={previewVersion.columns.map(formatDimensionName)}
            valueFields={previewVersion.values}
            valuesPlacement={previewVersion.valuesPlacement}
          />
        </Modal>
      )}
    </AppShell>
  );
}
