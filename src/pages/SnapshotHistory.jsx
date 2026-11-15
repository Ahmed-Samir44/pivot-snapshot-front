import { useEffect, useRef, useState } from "react";
import { useNavigate, useSearchParams } from "react-router-dom";
import { CalendarClock, ChevronRight, CirclePause, CirclePlay, Eye, FileSpreadsheet, History, Inbox, LayoutGrid, Link2, Loader2, MousePointerClick, Pencil, Plus, Search, Trash2, X } from "lucide-react";
import AppShell from "../components/layout/AppShell";
import Modal from "../components/forms/Modal";
import ConfirmModal from "../components/forms/ConfirmModal";
import ButtonProgressBar from "../components/forms/ButtonProgressBar";
import Popover from "../components/forms/Popover";
import KpiPicker from "../components/forms/KpiPicker";
import Toasts from "../components/layout/Toasts";
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
  linkKpi,
  getVersionByKpi,
  setSchedule,
} from "../services/snapshotApi";
import { getSchedulerStatus, connectScheduler, disconnectScheduler } from "../services/schedulerApi";
import { formatDimensionName, fieldCaptions } from "../utils/cubeMeta";
import { hideMeasuresLabelIfSingleValue } from "../utils/measuresPlacement";
import { DEFAULT_PIVOT_OPTIONS } from "../utils/pivotOptions";
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

// A saved Columns drilldown (C4): its columns carry `drill`, whose level is the slot in Columns.
function columnsDrillSlot(version) {
  return version.result.columnHeaders.find((h) => h.drill)?.drill.level ?? -1;
}

// What the preview grid needs to lay a saved drilldown out (read-only — no expansions to toggle).
function previewDrilldown(version) {
  if (isDrilldownVersion(version)) return { expandedMemberUniqueNames: [] };
  const slot = combinedDrillSlot(version);
  return slot >= 0 ? { topLevelField: version.rows[slot], expandedMemberUniqueNames: [] } : null;
}

// A version saved before pivot_sizebytes existed has no value for it (see
// DataverseSnapshotStorageService's own WHY comment) — shown as nothing rather than "0 B", which
// would misleadingly imply the version is actually empty.
function formatBytes(bytes) {
  if (!bytes) return null;
  const units = ["B", "KB", "MB", "GB"];
  let value = bytes;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${unitIndex === 0 ? value : value.toFixed(1)} ${units[unitIndex]}`;
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

// The column caption row's field names — a field's Custom Name (Field Settings) wins, as in the Builder.
function columnCaptions(version, fieldLayouts = []) {
  const captionOf = versionCaptions(version);
  return hideMeasuresLabelIfSingleValue(version.columns, version.values.length).map(
    (id) => (fieldLayouts ?? []).find((f) => f.field === id)?.customName || captionOf(id),
  );
}

// A saved version's field captions, like the Builder's ("BU Short Name"; two same-named fields keep
// their full names) — see fieldCaptions.
function versionCaptions(version) {
  const filters = parseBuilderState(version.builderStateJson)?.request?.filters ?? [];
  return fieldCaptions([...version.rows, ...version.columns, ...filters.map((f) => f.field)], formatDimensionName);
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

// A small tinted tag — a pivot's version count — instead of plain gray text next to the title,
// matching the rest of the page's badge-forward look (requested live: "يبقى بريميم").
// Schedule picker (2026-11-15) — PivotSchedulerBackgroundService.ParseHours' own frontend mirror:
// pivot_scheduleinterval is a free-form "every N hours" number as a string (old Hourly/Daily/
// Weekly values from before the user asked for "يقدر يحدد الـinterval بإيده" — a hand-typed custom
// interval — still read the same way, just not offered as a label a new pick would ever produce).
function parseScheduleHours(interval) {
  if (!interval || interval === "Off") return null;
  if (interval === "Hourly") return 1;
  if (interval === "Daily") return 24;
  if (interval === "Weekly") return 168;
  const hours = Number(interval);
  return Number.isFinite(hours) && hours > 0 ? hours : null;
}

function formatScheduleHours(hours) {
  if (hours === 1) return "Hourly";
  if (hours === 24) return "Daily";
  if (hours === 168) return "Weekly";
  const trimmed = Number(hours.toFixed(2)).toString();
  return `Every ${trimmed}h`;
}

// A small Popover (not an always-open control) so typing a custom number doesn't fight for space
// in the row's already-tight action icons — same pattern as the KPI-link Popover right next to it.
function ScheduleButton({ definition, onSetSchedule }) {
  const hours = parseScheduleHours(definition.scheduleInterval);
  const [draft, setDraft] = useState("");
  return (
    <Popover
      align="right"
      trigger={(toggle) => (
        <button
          type="button"
          onClick={(e) => {
            setDraft(hours ? String(hours) : "");
            toggle(e);
          }}
          className={`inline-flex items-center gap-1 rounded-full px-2 py-1.5 text-xs font-medium transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-primary/40 ${
            hours ? "text-primary-strong hover:bg-primary-50" : "text-muted hover:bg-slate-100"
          }`}
          title="Auto-run this pivot on a schedule"
          aria-label={`Schedule for pivot ${definition.name}`}
        >
          <CalendarClock className="h-4 w-4" />
          {hours ? formatScheduleHours(hours) : "Schedule: Off"}
        </button>
      )}
    >
      {(close) => (
        <div className="flex w-56 flex-col gap-2 text-xs">
          <p className="font-semibold text-ink">Auto-run every…</p>
          <div className="flex items-center gap-1.5">
            <input
              type="number"
              min="0"
              step="0.5"
              value={draft}
              onChange={(e) => setDraft(e.target.value)}
              placeholder="Off"
              className="input-field w-20 py-1.5"
              aria-label="Hours between runs"
            />
            <span className="text-muted">hours</span>
          </div>
          <div className="flex gap-1.5">
            {[
              ["Hourly", 1],
              ["Daily", 24],
              ["Weekly", 168],
            ].map(([label, value]) => (
              <button
                key={label}
                type="button"
                onClick={() => setDraft(String(value))}
                className="rounded-md border border-slate-200 px-2 py-1 text-[11px] text-ink hover:border-primary hover:text-primary-strong"
              >
                {label}
              </button>
            ))}
          </div>
          <div className="mt-1 flex justify-end gap-2">
            <button type="button" onClick={close} className="btn-secondary px-2 py-1 text-xs">
              Cancel
            </button>
            <button
              type="button"
              onClick={() => {
                const parsed = Number(draft);
                onSetSchedule(draft.trim() === "" || !Number.isFinite(parsed) || parsed <= 0 ? "Off" : String(parsed));
                close();
              }}
              className="btn-primary px-2 py-1 text-xs"
            >
              Apply
            </button>
          </div>
        </div>
      )}
    </Popover>
  );
}

function Pill({ children }) {
  return <span className="shrink-0 rounded-full bg-slate-100 px-2 py-0.5 text-[11px] font-semibold text-muted">{children}</span>;
}

// One row — used for pivots and versions alike, both Active and Deactivated lists — a bordered
// card per row (not just a hover tint) so rows read as distinct items instead of a dense block of
// text and icons running into each other (caught live 2026-09-23: "الشكل كله على بعضه"). Two lines
// — the title (+ an optional badge, e.g. a version count) on top, a muted detail line
// (a date, typically) below it — rather than squeezing both onto one line, for the same "premium"
// reason as Pill above.
function ListRow({ label, badge, meta, onClick, trailingIcon: TrailingIcon, actions, highlighted = false, checked, onToggleCheck, loading = false }) {
  return (
    <li
      className={`flex items-center gap-1 rounded-xl border py-1 pr-2 transition-all duration-150 ease-settle ${
        // A left accent bar (not just a tint) so the active pivot reads as distinct at a glance
        // instead of blending into its own hover-adjacent siblings — caught against a live
        // screenshot where the previous bg-only tint was easy to miss at a quick scan. A soft
        // elevated shadow (not just color) on both the selected row and on hover is what reads as
        // "premium" rather than flat (requested live: "عايز هم يبقوا عمودين بس الشكل يتغير يبقى
        // بريميم") — a plain colored border alone looked cheap next to the rest of the Synapse
        // design system's own elevated shadows (shadow-ds-sm/-lg, already used for cards/buttons).
        highlighted
          ? "border-primary/30 border-l-4 border-l-primary bg-primary-50 pl-1 shadow-ds-sm"
          : "border-slate-100 pl-2 hover:border-slate-200 hover:shadow-ds-sm"
      }`}
    >
      <label className="flex shrink-0 items-center justify-center rounded-lg p-1.5 hover:bg-slate-100">
        <input
          type="checkbox"
          checked={checked}
          onChange={onToggleCheck}
          className="h-4 w-4 accent-primary focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-primary/50"
          aria-label={`Select ${label}`}
        />
      </label>
      <button
        type="button"
        onClick={onClick}
        disabled={!onClick || loading}
        className={`flex flex-1 items-center justify-between gap-2 rounded-lg px-3 py-2.5 text-left transition-colors disabled:cursor-default disabled:hover:bg-transparent focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-primary/40 ${
          highlighted ? "text-primary-strong" : "hover:bg-slate-50"
        }`}
      >
        <span className="flex min-w-0 flex-col items-start gap-0.5">
          <span className="flex items-center gap-2">
            <span className={`font-medium ${highlighted ? "text-primary-strong" : "text-ink"}`}>{label}</span>
            {badge}
          </span>
          {meta && <span className="text-xs text-muted">{meta}</span>}
        </span>
        {/* Loading replaces the trailing icon entirely rather than sitting beside it — the eye icon
            itself means "view this," so leaving it visible next to a spinner while the click is
            already disabled (below) still reads as "nothing happened yet," which is exactly the
            ambiguity that let a slow preview fetch get clicked several times over (requested live
            2026-10-12: "بتاخد وقت طويل... بيخليها تتفتح كذا مرة" — no loading affordance at all
            meant every repeat click fired its own getVersion() call). */}
        {loading ? (
          <Loader2 className="h-4 w-4 shrink-0 animate-spin text-primary-strong" aria-hidden="true" />
        ) : (
          TrailingIcon && <TrailingIcon className="h-4 w-4 shrink-0 text-muted" aria-hidden="true" />
        )}
      </button>
      {actions && <div className="ml-1 flex shrink-0 items-center gap-0.5 border-l border-slate-100 pl-2">{actions}</div>}
    </li>
  );
}

// A dashed, centered placeholder for an empty list — reused for all three empty states on this
// page (no snapshots at all, no pivot picked yet, a pivot with zero versions) so they read as one
// deliberate pattern instead of three different plain-text strings.
function EmptyState({ icon: Icon, title, message }) {
  return (
    <li className="flex flex-col items-center gap-2 rounded-xl border border-dashed border-slate-200 px-4 py-10 text-center">
      <Icon className="h-7 w-7 text-slate-300" aria-hidden="true" />
      <p className="text-sm font-medium text-ink">{title}</p>
      {message && <p className="max-w-[22rem] text-xs text-muted">{message}</p>}
    </li>
  );
}

// Tinted pill instead of a plain dot + uppercase text — a flat label line read as an afterthought
// next to the rest of the page's new elevated-shadow rows (requested live: "يبقى بريميم").
const SECTION_TONES = {
  active: "bg-green-50 text-green-700",
  deactivated: "bg-orange-50 text-orange-700",
};
const SECTION_DOTS = {
  active: "bg-green-500",
  deactivated: "bg-orange-500",
};

function SectionLabel({ tone, children, action }) {
  return (
    <div className="mb-3 flex items-center justify-between gap-2">
      <span className={`inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-[11px] font-bold uppercase tracking-wide ${SECTION_TONES[tone]}`}>
        <span className={`h-1.5 w-1.5 rounded-full ${SECTION_DOTS[tone]}`} aria-hidden="true" />
        {children}
      </span>
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
  const [searchParams] = useSearchParams();
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
  // Why the last export is plain values instead of an Excel PivotTable (see excelPivot.js), or null.
  const [exportNotice, setExportNotice] = useState(null);
  const [searchTerm, setSearchTerm] = useState("");

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

  // Deep link from KpiLinks.jsx ("/history?definitionId=…") — selects that pivot the instant its
  // id shows up in the just-fetched list, once only (a plain dependency on `definitions` would
  // re-select it every time the list refreshes afterward, e.g. after a schedule change, undoing
  // whatever pivot the user picked themselves in the meantime).
  const appliedDeepLinkRef = useRef(false);
  useEffect(() => {
    if (appliedDeepLinkRef.current) return;
    const definitionId = searchParams.get("definitionId");
    if (!definitionId || !definitions.some((d) => d.id === definitionId)) return;
    appliedDeepLinkRef.current = true;
    selectDefinition(definitionId);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [definitions]);

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

  // Looked up from either list (not just the active one) — a pivot selected right before it gets
  // deactivated should still show its own name here, not blank out mid-view.
  const selectedPivotName = [...definitions, ...deactivatedDefinitions].find((d) => d.id === selectedDefinitionId)?.name ?? null;

  const matchesSearch = (d) => d.name.toLowerCase().includes(searchTerm.trim().toLowerCase());
  const filteredDefinitions = definitions.filter(matchesSearch);
  const filteredDeactivatedDefinitions = deactivatedDefinitions.filter(matchesSearch);
  const totalPivotCount = definitions.length + deactivatedDefinitions.length;
  const visiblePivotCount = filteredDefinitions.length + filteredDeactivatedDefinitions.length;

  const handleExport = async () => {
    if (!previewVersion) return;
    setExporting(true);
    setExportError(null);
    try {
      // Dynamic import — see PivotBuilder.jsx's handleExport for why (ExcelJS is ~1MB, only
      // worth loading for someone who actually clicks Export).
      const { exportPivotToExcel } = await import("../utils/exportExcel");
      const pivotName = definitions.find((d) => d.id === selectedDefinitionId)?.name ?? "Pivot";
      const exported = await exportPivotToExcel({
        result: previewVersion.result,
        rows: previewVersion.rows,
        columns: previewVersion.columns,
        rowFieldLabels: previewVersion.rows.map(versionCaptions(previewVersion)),
        columnFieldLabels: columnCaptions(previewVersion, parseBuilderState(previewVersion.builderStateJson)?.fieldLayouts),
        valueFields: previewVersion.values,
        filename: `${pivotName}_v${previewVersion.versionNumber}`,
        // Same layout as the preview grid above, so the export matches what's on screen.
        layout: parseBuilderState(previewVersion.builderStateJson)?.reportLayout ?? "Tabular",
        repeatItemLabels: parseBuilderState(previewVersion.builderStateJson)?.repeatItemLabels ?? false,
        fieldLayouts: parseBuilderState(previewVersion.builderStateJson)?.fieldLayouts ?? [],
        drilldown: isDrilldownVersion(previewVersion),
        drillSlot: combinedDrillSlot(previewVersion),
        columnsDrillSlot: columnsDrillSlot(previewVersion),
        drilldownSubtotals: parseBuilderState(previewVersion.builderStateJson)?.request?.subtotals ?? "Top",
        // PivotTable Options and the Report Filter area, from the saved builder state (versions
        // saved before 2026-10-18 have no options — Excel's defaults).
        pivotOptions: { ...DEFAULT_PIVOT_OPTIONS, ...(parseBuilderState(previewVersion.builderStateJson)?.pivotOptions ?? {}) },
        // Formats set on individual cells (Home › Number) when this version was saved.
        cellFormats: parseBuilderState(previewVersion.builderStateJson)?.cellFormats ?? null,
        filters: (parseBuilderState(previewVersion.builderStateJson)?.request?.filters ?? []).filter((f) => f.field),
        filterFieldLabels: Object.fromEntries((parseBuilderState(previewVersion.builderStateJson)?.request?.filters ?? []).map((f) => [f.field, versionCaptions(previewVersion)(f.field)])),
      });
      setExportNotice(exported.pivotTable ? null : `Exported as values, not a PivotTable: ${exported.reason}.`);
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
  const applyEdit = (builderStateJson) => {
    try {
      localStorage.setItem(BUILDER_STATE_STORAGE_KEY, builderStateJson);
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
  // Holds the builderStateJson waiting on the overwrite confirmation (null = nothing pending) —
  // not just a boolean, since Edit can now be triggered straight from a version ROW (below), not
  // only from the already-open preview, so there's no single "the current version" to fall back to.
  const [pendingEditOverwrite, setPendingEditOverwrite] = useState(null);
  const handleEdit = (builderStateJson) => {
    if (!builderStateJson) return;
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
      setPendingEditOverwrite(builderStateJson);
      return;
    }
    applyEdit(builderStateJson);
  };

  // KPI linking (2026-11-12, "عايز اقدر اربط kpi بال snapshot... بعد كدا اختار kpi معين يتبنى ال
  // pivot table") — link this saved pivot to a KPI, and separately pick a KPI to load whatever
  // pivot is already linked to it straight into the Builder, reusing applyEdit/handleEdit exactly
  // like the row-level Edit button above (same localStorage-write-then-navigate restore path).
  const [kpiLinkMessage, setKpiLinkMessage] = useState(null);
  const [kpiLinkError, setKpiLinkError] = useState(null);

  // Scheduling (2026-11-15) — whether a scheduled-run credential is currently stored server-side
  // (see ScheduledRunTokenStore), fetched once on mount. null while loading, so the button doesn't
  // flash "Enable Scheduling" for an instant before showing "Scheduling enabled".
  const [schedulerConnected, setSchedulerConnected] = useState(null);
  const [schedulerBusy, setSchedulerBusy] = useState(false);
  const [schedulerError, setSchedulerError] = useState(null);
  useEffect(() => {
    getSchedulerStatus()
      .then((status) => setSchedulerConnected(status.connected))
      .catch(() => setSchedulerConnected(false));
  }, []);
  const handleToggleScheduler = async () => {
    setSchedulerBusy(true);
    setSchedulerError(null);
    try {
      if (schedulerConnected) {
        await disconnectScheduler();
        setSchedulerConnected(false);
      } else {
        await connectScheduler();
        setSchedulerConnected(true);
      }
    } catch (err) {
      setSchedulerError(err.message);
    } finally {
      setSchedulerBusy(false);
    }
  };
  const [scheduleError, setScheduleError] = useState(null);
  const handleSetSchedule = async (definitionId, interval) => {
    try {
      await setSchedule(definitionId, interval);
      await refreshDefinitionLists();
    } catch (err) {
      setScheduleError(err.message);
    }
  };
  const handleLinkKpi = async (definitionId, definitionName, kpi, close) => {
    try {
      await linkKpi(definitionId, kpi.id);
      setKpiLinkMessage(`Linked "${definitionName}" to KPI "${kpi.name}".`);
      close();
    } catch (err) {
      setKpiLinkError(err.message);
    }
  };
  const handleOpenByKpi = async (kpi, close) => {
    try {
      const version = await getVersionByKpi(kpi.id);
      if (!version) {
        setKpiLinkError(`No pivot is linked to "${kpi.name}" yet.`);
        return;
      }
      close();
      handleEdit(version.builderStateJson);
    } catch (err) {
      setKpiLinkError(err.message);
    }
  };

  // Edit straight from a version's own row in the list (2026-11-12, "عايز زرار ال edit يظهر هنا")
  // — the list only ever carries PivotSnapshotVersionSummary (no builderStateJson, see
  // DataverseSnapshotStorageService's own two DTOs), so this fetches the full version first, same
  // as opening the preview does. Shares loadingVersionId with openPreview below (one fetch at a
  // time, same reasoning) so this row's spinner and preview's spinner never fight each other.
  const editVersionFromRow = async (versionId) => {
    if (loadingVersionId) return;
    setLoadingVersionId(versionId);
    setPreviewError(null);
    try {
      const version = await getVersion(versionId);
      if (!version.builderStateJson) {
        setPreviewError("This version was saved before Edit support existed.");
        return;
      }
      handleEdit(version.builderStateJson);
    } catch (err) {
      setPreviewError(err.message);
    } finally {
      setLoadingVersionId(null);
    }
  };

  return (
    <AppShell>
      <header className="mb-6 flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
        <div>
          <h1 className="text-3xl font-bold text-ink">Snapshot History</h1>
          <p className="mt-1 text-muted">Every saved pivot, and every version saved under it.</p>
        </div>
        <div className="flex shrink-0 items-center gap-2">
          <Popover
            align="right"
            trigger={(toggle) => (
              <button type="button" onClick={toggle} className="btn-secondary inline-flex items-center gap-1.5">
                <Link2 className="h-4 w-4" />
                Open by KPI…
              </button>
            )}
          >
            {(close) => <KpiPicker onSelect={(kpi) => handleOpenByKpi(kpi, close)} />}
          </Popover>
          {/* Scheduling (2026-11-15) — connects THIS signed-in account as the one scheduled runs
              authenticate as (see ScheduledRunTokenStore's own WHY comment), a bridge until this
              tenant has a real service identity for Dataverse. */}
          <button
            type="button"
            onClick={handleToggleScheduler}
            disabled={schedulerConnected === null || schedulerBusy}
            className={`inline-flex items-center gap-1.5 disabled:cursor-not-allowed disabled:opacity-50 ${schedulerConnected ? "btn-secondary !border-green-300 !text-green-700" : "btn-secondary"}`}
            title={schedulerConnected ? "Scheduled pivots run using your account — click to disconnect" : "Let scheduled pivots run automatically using your account"}
          >
            <CalendarClock className="h-4 w-4" />
            {schedulerBusy ? "…" : schedulerConnected ? "Scheduling enabled" : "Enable Scheduling"}
          </button>
          <button type="button" onClick={() => navigate("/")} className="btn-primary inline-flex items-center gap-1.5">
            <Plus className="h-4 w-4" />
            New Pivot
          </button>
        </div>
      </header>

      <div className="relative mb-6">
        <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted" aria-hidden="true" />
        <input
          type="text"
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
          placeholder="Search saved pivots…"
          className="w-full rounded-lg border border-slate-200 bg-white py-2 pl-9 pr-4 text-sm shadow-sm transition-colors focus:border-primary focus:outline-none focus:ring-2 focus:ring-primary/20"
        />
      </div>

      {error && <p className="mb-6 rounded-xl border border-red-300 bg-red-50 px-4 py-3 text-red-800">{error}</p>}

      <div className="grid grid-cols-1 gap-4 lg:grid-cols-2 lg:items-start lg:gap-6">
        <div className="rounded-2xl border border-slate-200/70 bg-white p-5 shadow-ds-lg">
          <div className="mb-5 flex items-center gap-3">
            <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-primary to-primary-deep text-white shadow-button">
              <LayoutGrid className="h-5 w-5" aria-hidden="true" />
            </span>
            <div>
              <h3 className="text-sm font-bold uppercase tracking-wide text-muted">Saved Pivots</h3>
              <p className="text-xs text-muted">{totalPivotCount} total</p>
            </div>
          </div>

          <SectionLabel tone="active">Active</SectionLabel>
          <BulkActionBar
            selection={activeDefSelection}
            onDeactivate={() => {
              const ids = [...activeDefSelection.ids];
              const versionCount = definitions.filter((d) => ids.includes(d.id)).reduce((sum, d) => sum + d.versionCount, 0);
              setPendingConfirm({ action: "deactivate-definition", ids, label: countLabel(ids.length, "pivot"), versionCount });
            }}
          />
          <ul className="space-y-2">
            {filteredDefinitions.map((d) => (
              <ListRow
                key={d.id}
                label={d.name}
                badge={
                  <>
                    <Pill>{d.versionCount} version{d.versionCount === 1 ? "" : "s"}</Pill>
                    {d.linkedKpiName && (
                      <Pill>
                        <Link2 className="-ml-0.5 mr-0.5 inline h-3 w-3 align-text-bottom" />
                        {d.linkedKpiName}
                      </Pill>
                    )}
                  </>
                }
                meta={`Created ${new Date(d.createdAtUtc).toLocaleDateString()}`}
                onClick={() => selectDefinition(d.id)}
                trailingIcon={ChevronRight}
                highlighted={selectedDefinitionId === d.id}
                checked={activeDefSelection.ids.has(d.id)}
                onToggleCheck={() => activeDefSelection.toggle(d.id)}
                actions={
                  <>
                    {/* Scheduling (2026-11-15) — only offered once this pivot is linked to a KPI
                        (see SetScheduleAsync's own WHY comment: no clear reason to auto-run one
                        that isn't feeding a KPI). */}
                    {d.linkedKpiName && (
                      <ScheduleButton definition={d} onSetSchedule={(interval) => handleSetSchedule(d.id, interval)} />
                    )}
                    <Popover
                      align="right"
                      trigger={(toggle) => (
                        <button
                          type="button"
                          onClick={toggle}
                          className="rounded-full p-2 text-primary-strong transition-colors hover:bg-primary-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-primary/40"
                          title="Link this pivot to a KPI"
                          aria-label={`Link pivot ${d.name} to a KPI`}
                        >
                          <Link2 className="h-4 w-4" />
                        </button>
                      )}
                    >
                      {(close) => <KpiPicker onSelect={(kpi) => handleLinkKpi(d.id, d.name, kpi, close)} />}
                    </Popover>
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
                      className="rounded-full p-2 text-orange-500 transition-colors hover:bg-orange-50 hover:text-orange-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-orange-300"
                      title="Deactivate this pivot"
                      aria-label={`Deactivate pivot ${d.name}`}
                    >
                      <CirclePause className="h-4 w-4" />
                    </button>
                  </>
                }
              />
            ))}
            {definitions.length === 0 && (
              <EmptyState icon={Inbox} title="No snapshots saved yet" message="Save a pivot from the Builder to see it here." />
            )}
            {definitions.length > 0 && filteredDefinitions.length === 0 && (
              <EmptyState icon={Search} title="No matches" message="Nothing matches your search." />
            )}
          </ul>

          {filteredDeactivatedDefinitions.length > 0 && (
            <div className="mt-6 border-t border-slate-100 pt-5">
              <SectionLabel tone="deactivated">Deactivated</SectionLabel>
              <BulkActionBar
                selection={deactivatedDefSelection}
                onReactivate={() => runDefinitionAction(reactivateDefinition, [...deactivatedDefSelection.ids])}
                onDelete={() => {
                  const ids = [...deactivatedDefSelection.ids];
                  setPendingConfirm({ action: "delete-definition", ids, label: countLabel(ids.length, "pivot") });
                }}
              />
              <ul className="space-y-2">
                {filteredDeactivatedDefinitions.map((d) => (
                  <ListRow
                    key={d.id}
                    label={d.name}
                    badge={
                  <>
                    <Pill>{d.versionCount} version{d.versionCount === 1 ? "" : "s"}</Pill>
                    {d.linkedKpiName && (
                      <Pill>
                        <Link2 className="-ml-0.5 mr-0.5 inline h-3 w-3 align-text-bottom" />
                        {d.linkedKpiName}
                      </Pill>
                    )}
                  </>
                }
                    meta={`Created ${new Date(d.createdAtUtc).toLocaleDateString()}`}
                    checked={deactivatedDefSelection.ids.has(d.id)}
                    onToggleCheck={() => deactivatedDefSelection.toggle(d.id)}
                    actions={
                      <>
                        <button
                          type="button"
                          onClick={() => runDefinitionAction(reactivateDefinition, [d.id])}
                          className="rounded-full p-2 text-green-600 transition-colors hover:bg-green-50 hover:text-green-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-green-300"
                          title="Reactivate this pivot"
                          aria-label={`Reactivate pivot ${d.name}`}
                        >
                          <CirclePlay className="h-4 w-4" />
                        </button>
                        <button
                          type="button"
                          onClick={() => setPendingConfirm({ action: "delete-definition", ids: [d.id], label: d.name })}
                          className="rounded-full p-2 text-muted transition-colors hover:bg-red-50 hover:text-red-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-red-300"
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

          <p className="mt-5 border-t border-slate-100 pt-3 text-xs text-muted">
            Showing {visiblePivotCount} of {totalPivotCount} pivot{totalPivotCount === 1 ? "" : "s"}
          </p>
        </div>

        <div className="rounded-2xl border border-slate-200/70 bg-white p-5 shadow-ds-lg">
          {/* Names which pivot's versions are below — the left column's selection and this
              column's contents used to have no visible link between them (caught against a live
              screenshot: selecting a pivot populated the list with no indication WHICH pivot it
              belonged to once you'd scrolled or come back to the tab later). */}
          <div className="mb-5 flex items-start justify-between gap-2">
            <div className="flex items-center gap-3">
              {/* Dark/ink instead of the Pivots card's teal — the two badges read as two distinct
                  sections of the same system at a glance, not an accidental repeat. */}
              <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-ink text-white shadow-button">
                <History className="h-5 w-5" aria-hidden="true" />
              </span>
              <div>
                <h3 className="text-sm font-bold uppercase tracking-wide text-muted">Versions</h3>
                {selectedPivotName ? (
                  <p className="mt-0.5 text-base font-semibold text-ink">{selectedPivotName}</p>
                ) : (
                  <p className="mt-0.5 text-sm text-muted">No pivot selected</p>
                )}
              </div>
            </div>
            {selectedDefinitionId && (
              <button
                type="button"
                onClick={() => {
                  setSelectedDefinitionId(null);
                  setVersions([]);
                  setDeactivatedVersions([]);
                }}
                className="shrink-0 rounded-lg px-2 py-1 text-xs font-semibold text-muted transition-colors hover:bg-slate-50 hover:text-ink focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-primary/40"
              >
                Clear Selection
              </button>
            )}
          </div>
          {previewError && <p className="mb-3 rounded-xl border border-red-300 bg-red-50 px-3 py-2 text-sm text-red-800">{previewError}</p>}

          <SectionLabel tone="active">Active</SectionLabel>
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
                meta={[new Date(v.createdAtUtc).toLocaleString(), formatBytes(v.sizeBytes)].filter(Boolean).join(" · ")}
                onClick={() => openPreview(v.id, v.versionNumber)}
                trailingIcon={Eye}
                loading={loadingVersionId === v.id}
                checked={activeVersionSelection.ids.has(v.id)}
                onToggleCheck={() => activeVersionSelection.toggle(v.id)}
                actions={
                  <>
                    <button
                      type="button"
                      onClick={() => editVersionFromRow(v.id)}
                      disabled={Boolean(loadingVersionId)}
                      className="rounded-full p-2 text-primary-strong transition-colors hover:bg-primary-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-primary/40 disabled:cursor-not-allowed disabled:opacity-50"
                      title="Edit this version in the Builder"
                      aria-label={`Edit version ${v.versionNumber}`}
                    >
                      <Pencil className="h-4 w-4" />
                    </button>
                    <button
                      type="button"
                      onClick={() => setPendingConfirm({ action: "deactivate-version", ids: [v.id], label: versionLabel(v) })}
                      className="rounded-full p-2 text-orange-500 transition-colors hover:bg-orange-50 hover:text-orange-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-orange-300"
                      title="Deactivate this version"
                      aria-label={`Deactivate version ${v.versionNumber}`}
                    >
                      <CirclePause className="h-4 w-4" />
                    </button>
                  </>
                }
              />
            ))}
            {!selectedDefinitionId && (
              <EmptyState icon={MousePointerClick} title="Pick a saved pivot" message="Select one on the left to see its versions here." />
            )}
            {selectedDefinitionId && versions.length === 0 && (
              <EmptyState icon={History} title="No active versions" message="Save a version of this pivot from the Builder to see it here." />
            )}
          </ul>

          {selectedDefinitionId && deactivatedVersions.length > 0 && (
            <div className="mt-6 border-t border-slate-100 pt-5">
              <SectionLabel
                tone="deactivated"
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
                    meta={[new Date(v.createdAtUtc).toLocaleString(), formatBytes(v.sizeBytes)].filter(Boolean).join(" · ")}
                    onClick={() => openPreview(v.id, v.versionNumber)}
                    trailingIcon={Eye}
                    loading={loadingVersionId === v.id}
                    checked={deactivatedVersionSelection.ids.has(v.id)}
                    onToggleCheck={() => deactivatedVersionSelection.toggle(v.id)}
                    actions={
                      <>
                        <button
                          type="button"
                          onClick={() => editVersionFromRow(v.id)}
                          disabled={Boolean(loadingVersionId)}
                          className="rounded-full p-2 text-primary-strong transition-colors hover:bg-primary-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-primary/40 disabled:cursor-not-allowed disabled:opacity-50"
                          title="Edit this version in the Builder"
                          aria-label={`Edit version ${v.versionNumber}`}
                        >
                          <Pencil className="h-4 w-4" />
                        </button>
                        <button
                          type="button"
                          onClick={() => runVersionAction(reactivateVersion, [v.id])}
                          className="rounded-full p-2 text-green-600 transition-colors hover:bg-green-50 hover:text-green-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-green-300"
                          title="Reactivate this version"
                          aria-label={`Reactivate version ${v.versionNumber}`}
                        >
                          <CirclePlay className="h-4 w-4" />
                        </button>
                        <button
                          type="button"
                          onClick={() => setPendingConfirm({ action: "delete-version", ids: [v.id], label: versionLabel(v) })}
                          className="rounded-full p-2 text-muted transition-colors hover:bg-red-50 hover:text-red-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-red-300"
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
                onClick={() => handleEdit(previewVersion.builderStateJson)}
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
                className="btn-secondary relative overflow-hidden inline-flex items-center gap-1.5 disabled:cursor-not-allowed disabled:opacity-50"
              >
                <FileSpreadsheet className="h-4 w-4" style={{ color: "#217346" }} />
                {exporting ? "Exporting…" : "Export to Excel"}
                <ButtonProgressBar active={exporting} tone="dark" />
              </button>
              {exportError && <span className="text-sm font-medium text-red-700">{exportError}</span>}
              {exportNotice && <span className="text-xs text-amber-700">{exportNotice}</span>}
            </div>
            <PivotGrid
              cellFormats={parseBuilderState(previewVersion.builderStateJson)?.cellFormats ?? null}
              result={previewVersion.result}
              rows={previewVersion.rows}
              columns={previewVersion.columns}
              rowFieldLabels={previewVersion.rows.map(versionCaptions(previewVersion))}
              columnFieldLabels={columnCaptions(previewVersion, builderState?.fieldLayouts)}
              valueFields={previewVersion.values}
              layout={builderState?.reportLayout ?? "Tabular"}
              repeatItemLabels={builderState?.repeatItemLabels ?? false}
              fieldLayouts={builderState?.fieldLayouts ?? []}
              // Read-only here: no onToggleDrilldownMember, so no +/- toggles — just the frozen rows.
              rowsHierarchyDrilldown={previewDrilldown(previewVersion)}
              columnsDrillSlot={columnsDrillSlot(previewVersion)}
              // Older versions (no builder state) predate drilldown subtotals — Top = how they looked.
              drilldownSubtotals={builderState?.request?.subtotals ?? "Top"}
              boundedHeight={false}
              pivotOptions={{ ...DEFAULT_PIVOT_OPTIONS, ...(builderState?.pivotOptions ?? {}) }}
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
            const builderStateJson = pendingEditOverwrite;
            setPendingEditOverwrite(null);
            applyEdit(builderStateJson);
          }}
          onCancel={() => setPendingEditOverwrite(null)}
        />
      )}
      <Toasts
        items={[
          kpiLinkMessage && { tone: "success", text: kpiLinkMessage, onClose: () => setKpiLinkMessage(null) },
          kpiLinkError && { tone: "error", text: kpiLinkError, onClose: () => setKpiLinkError(null) },
          schedulerError && { tone: "error", text: schedulerError, onClose: () => setSchedulerError(null) },
          scheduleError && { tone: "error", text: scheduleError, onClose: () => setScheduleError(null) },
        ].filter(Boolean)}
      />
    </AppShell>
  );
}
