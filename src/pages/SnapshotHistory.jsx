import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { ChevronRight, CirclePause, CirclePlay, Eye, FileSpreadsheet, Loader2, Pencil, Trash2, X } from "lucide-react";
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
import { hideMeasuresLabelIfSingleValue } from "../utils/measuresPlacement";
import { BUILDER_STATE_STORAGE_KEY } from "./PivotBuilder";

// BuilderStateJson is null for any version saved BEFORE this existed (2026-10-07) — parsed once
// here rather than at each of its 3 call sites (the preview's layout/repeatItemLabels props, and
// the Edit button's own enabled check). Falls back to the same defaults PivotGrid/PivotBuilder
// already use on their own, so an old version's preview renders exactly as it always has.
// "Version 3" normally, "Version 3 — Q3 Final" when a label was given at save time (see
// PivotBuilder.jsx's versionName field) — additive, never replacing the number itself.
function versionLabel(v) {
  return v.versionName ? `Version ${v.versionNumber} — ${v.versionName}` : `Version ${v.versionNumber}`;
}

// A saved per-member drilldown pivot (PLAN 8) has no Rows fields but does have row labels — its
// hierarchy replaces Rows entirely. Read off the version itself (not its builder state, which an
// older version may lack), so the preview and export draw it as one indented "Row Labels" column.
function isDrilldownVersion(version) {
  return version.rows.length === 0 && version.result.rowHeaders.some((h) => h.labels.length > 0);
}

// A saved drilldown COMBINED with other Rows fields (Drilldown v3): its rows carry `drill` info,
// whose level is the slot's position in Rows. -1 = not one.
function combinedDrillSlot(version) {
  return version.result.rowHeaders.find((h) => h.drill)?.drill.level ?? -1;
}

// What the preview grid needs to lay a saved drilldown out (read-only — no expansions to toggle).
function previewDrilldown(version) {
  if (isDrilldownVersion(version)) return { expandedMemberUniqueNames: [] };
  const slot = combinedDrillSlot(version);
  return slot >= 0 ? { topLevelField: version.rows[slot], expandedMemberUniqueNames: [] } : null;
}

function parseBuilderState(builderStateJson) {
  if (!builderStateJson) return null;
  try {
    return JSON.parse(builderStateJson);
  } catch {
    // Treated the same as "never existed" — a corrupted blob shouldn't crash the preview, just
    // fall back to defaults and disable Edit (see its own disabled check below).
    return null;
  }
}

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
function ListRow({ label, sublabel, onClick, trailingIcon: TrailingIcon, actions, highlighted = false, checked, onToggleCheck, loading = false }) {
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
        disabled={!onClick || loading}
        className={`flex flex-1 items-center justify-between gap-2 rounded-lg px-3 py-2.5 text-left disabled:cursor-default disabled:hover:bg-transparent ${
          highlighted ? "bg-gold/20 font-semibold" : "hover:bg-slate-50"
        }`}
      >
        <span>
          <span className="font-medium text-ink">{label}</span>
          {sublabel && <span className="ml-2 text-xs text-muted">{sublabel}</span>}
        </span>
        {/* Loading replaces the trailing icon entirely rather than sitting beside it — the eye icon
            itself means "view this," so leaving it visible next to a spinner while the click is
            already disabled (below) still reads as "nothing happened yet," which is exactly the
            ambiguity that let a slow preview fetch get clicked several times over (requested live
            2026-10-12: "بتاخد وقت طويل... بيخليها تتفتح كذا مرة" — no loading affordance at all
            meant every repeat click fired its own getVersion() call). */}
        {loading ? (
          <Loader2 className="h-4 w-4 shrink-0 animate-spin text-gold-strong" aria-hidden="true" />
        ) : (
          TrailingIcon && <TrailingIcon className="h-4 w-4 shrink-0 text-muted" aria-hidden="true" />
        )}
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
  const navigate = useNavigate();
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
  // Which version's getVersion() fetch is currently in flight (its full result/HTML payload can be
  // a slow Dataverse round-trip) — drives both the spinner on that row (ListRow's own `loading`
  // prop) and re-click protection in openPreview below (see its own WHY comment).
  const [loadingVersionId, setLoadingVersionId] = useState(null);
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
    // Re-click guard — ListRow already disables the clicked row's own button while its spinner
    // shows, but that alone doesn't stop a DIFFERENT row being clicked mid-fetch, or a double-click
    // landing both calls before the first re-render disables anything. One fetch at a time, full
    // stop, is the simplest correct rule here — a second click just gets dropped, not queued.
    if (loadingVersionId) return;
    setLoadingVersionId(versionId);
    setPreviewError(null);
    setExportError(null);
    try {
      const version = await getVersion(versionId);
      setPreviewVersion({ ...version, versionNumber });
    } catch (err) {
      setPreviewError(err.message);
    } finally {
      setLoadingVersionId(null);
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
        rows: previewVersion.rows,
        columns: previewVersion.columns,
        rowFieldLabels: previewVersion.rows.map(formatDimensionName),
        columnFieldLabels: hideMeasuresLabelIfSingleValue(previewVersion.columns, previewVersion.values.length).map(formatDimensionName),
        valueFields: previewVersion.values,
        filename: `${pivotName}_v${previewVersion.versionNumber}`,
        // Same layout as the preview grid above, so the export matches what's on screen.
        layout: parseBuilderState(previewVersion.builderStateJson)?.reportLayout ?? "Tabular",
        repeatItemLabels: parseBuilderState(previewVersion.builderStateJson)?.repeatItemLabels ?? false,
        fieldLayouts: parseBuilderState(previewVersion.builderStateJson)?.fieldLayouts ?? [],
        drilldown: isDrilldownVersion(previewVersion),
        drillSlot: combinedDrillSlot(previewVersion),
        drilldownSubtotals: parseBuilderState(previewVersion.builderStateJson)?.request?.subtotals ?? "Top",
      });
    } catch (err) {
      setExportError(err.message);
    } finally {
      setExporting(false);
    }
  };

  // Hands this version's COMPLETE saved config to the Builder page via the exact same localStorage
  // slot it already restores from on its own mount (see PivotBuilder.jsx's BUILDER_STATE_STORAGE_KEY
  // WHY comment) — Edit is just "restore THIS instead of whatever was there before," reusing that
  // existing restore path rather than a new one. Only offered when builderStateJson actually exists
  // (see its own disabled check at the button below) — a version saved before 2026-10-07 has no way
  // to recover its Sort/Totals/Calculated Items/Groupings, and silently dropping those on Edit would
  // be a worse surprise than just not offering it.
  const applyEdit = () => {
    try {
      localStorage.setItem(BUILDER_STATE_STORAGE_KEY, previewVersion.builderStateJson);
    } catch {
      // Same no-op-on-failure as every other localStorage write in this app — worst case, the
      // Builder just opens blank instead of restoring this version, no worse than Edit not existing.
    }
    navigate("/");
  };

  // Edit overwrites whatever's currently in that SAME localStorage slot — if the Builder has an
  // in-progress, un-saved pivot sitting there, Edit would silently discard it with no way back
  // (caught live 2026-10-11, right after shipping Edit itself: "عندك شغل مش محفوظ، متأكد؟"). Only
  // warns when there's something actually worth losing (a non-empty Rows/Columns/Values, same
  // "isPivotEmpty" definition PivotBuilder.jsx's own auto-run effect uses) — an untouched/already-
  // reset Builder has nothing to warn about.
  const [pendingEditOverwrite, setPendingEditOverwrite] = useState(false);
  const handleEdit = () => {
    if (!previewVersion?.builderStateJson) return;
    let hasUnsavedWork = false;
    try {
      const existing = JSON.parse(localStorage.getItem(BUILDER_STATE_STORAGE_KEY) || "null");
      const req = existing?.request;
      hasUnsavedWork = Boolean(req && (req.rows?.length > 0 || req.columns?.length > 0 || req.values?.length > 0));
    } catch {
      // An unreadable/corrupted existing entry has nothing worth protecting either.
      hasUnsavedWork = false;
    }
    if (hasUnsavedWork) {
      setPendingEditOverwrite(true);
      return;
    }
    applyEdit();
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
                label={versionLabel(v)}
                sublabel={new Date(v.createdAtUtc).toLocaleString()}
                onClick={() => openPreview(v.id, v.versionNumber)}
                trailingIcon={Eye}
                loading={loadingVersionId === v.id}
                checked={activeVersionSelection.ids.has(v.id)}
                onToggleCheck={() => activeVersionSelection.toggle(v.id)}
                actions={
                  <button
                    type="button"
                    onClick={() => setPendingConfirm({ action: "deactivate-version", ids: [v.id], label: versionLabel(v) })}
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
                    label={versionLabel(v)}
                    sublabel={new Date(v.createdAtUtc).toLocaleString()}
                    onClick={() => openPreview(v.id, v.versionNumber)}
                    trailingIcon={Eye}
                    loading={loadingVersionId === v.id}
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
                          onClick={() => setPendingConfirm({ action: "delete-version", ids: [v.id], label: versionLabel(v) })}
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

      {previewVersion && (() => {
        const builderState = parseBuilderState(previewVersion.builderStateJson);
        return (
          <Modal title={`${versionLabel(previewVersion)} — interactive preview`} onClose={() => setPreviewVersion(null)}>
            <div className="mb-3 flex flex-wrap items-center gap-3">
              <button
                type="button"
                onClick={handleEdit}
                disabled={!builderState}
                title={builderState ? "Open this version in the Builder to keep editing it" : "This version was saved before Edit support existed"}
                className="btn-secondary inline-flex items-center gap-1.5 disabled:cursor-not-allowed disabled:opacity-50"
              >
                <Pencil className="h-4 w-4" />
                Edit
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
              {exportError && <span className="text-sm font-medium text-red-700">{exportError}</span>}
            </div>
            <PivotGrid
              result={previewVersion.result}
              rows={previewVersion.rows}
              columns={previewVersion.columns}
              rowFieldLabels={previewVersion.rows.map(formatDimensionName)}
              columnFieldLabels={hideMeasuresLabelIfSingleValue(previewVersion.columns, previewVersion.values.length).map(formatDimensionName)}
              valueFields={previewVersion.values}
              layout={builderState?.reportLayout ?? "Tabular"}
              repeatItemLabels={builderState?.repeatItemLabels ?? false}
              fieldLayouts={builderState?.fieldLayouts ?? []}
              // Read-only here: no onToggleDrilldownMember, so no +/- toggles — just the frozen rows.
              rowsHierarchyDrilldown={previewDrilldown(previewVersion)}
              // Older versions (no builder state) predate drilldown subtotals — Top = how they looked.
              drilldownSubtotals={builderState?.request?.subtotals ?? "Top"}
              boundedHeight={false}
            />
          </Modal>
        );
      })()}

      {/* Rendered AFTER previewVersion's own Modal above (not before) — with both using the
          same z-50 backdrop, DOM order decides which one visually sits on top of the other
          (later wins). This needs to win, since it opens FROM a button inside that same preview
          Modal — placed earlier, it rendered BEHIND the still-open preview, invisible underneath
          it (caught live 2026-10-11, "الصفحة اللي دوست فيها edit مش بتتقفل فال overrite دي بتظهر
          تحتها"). */}
      {pendingEditOverwrite && (
        <ConfirmModal
          title="Overwrite unsaved work?"
          message="The Builder page has an in-progress pivot that hasn't been saved as a snapshot yet. Opening this version for editing will replace it — that unsaved work can't be recovered afterward."
          confirmLabel="Overwrite and continue"
          danger
          onConfirm={() => {
            setPendingEditOverwrite(false);
            applyEdit();
          }}
          onCancel={() => setPendingEditOverwrite(false)}
        />
      )}
    </AppShell>
  );
}
