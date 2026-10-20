import { Loader2, Minus, PanelRightClose, PanelRightOpen, Plus } from "lucide-react";
import { ZOOM_LEVELS } from "../../utils/zoom";


const statFormat = new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 });

// Excel's status bar along the bottom of the workspace (2026-10-20 UI refresh): the calculation
// state on the left, Average / Count / Sum of the selected cells (shown, like Excel, once 2+ cells
// are selected — a single cell already shows its own value), then the view controls (Field List
// toggle and zoom) on the right.
export default function StatusBar({ loading, pendingLayout, stats, zoom, onZoomChange, fieldListShown, onToggleFieldList }) {
  const zoomIndex = ZOOM_LEVELS.indexOf(zoom);
  const stepZoom = (delta) => {
    const index = zoomIndex >= 0 ? zoomIndex : ZOOM_LEVELS.indexOf(1);
    onZoomChange(ZOOM_LEVELS[Math.min(ZOOM_LEVELS.length - 1, Math.max(0, index + delta))]);
  };
  const showStats = stats && stats.cellCount >= 2 && stats.count > 0;

  return (
    <footer className="flex h-8 shrink-0 items-center gap-4 rounded-md border border-slate-200 bg-slate-50 px-3 text-xs text-slate-600" aria-label="Status bar">
      <span className="flex items-center gap-1.5" role="status">
        {loading ? (
          <>
            <Loader2 className="h-3.5 w-3.5 animate-spin text-gold-strong" />
            Updating…
          </>
        ) : pendingLayout ? (
          <span className="text-amber-700">Layout changes pending — Update to apply</span>
        ) : (
          "Ready"
        )}
      </span>

      <div className="ml-auto flex items-center gap-4">
        {showStats && (
          <dl className="flex items-center gap-4 tabular-nums" aria-label="Selection summary">
            {stats.average !== null && <Stat label="Average" value={statFormat.format(stats.average)} />}
            <Stat label="Count" value={statFormat.format(stats.count)} />
            {stats.sum !== null && <Stat label="Sum" value={statFormat.format(stats.sum)} />}
          </dl>
        )}

        <span className="h-4 w-px bg-slate-200" aria-hidden="true" />

        <button
          type="button"
          onClick={onToggleFieldList}
          aria-pressed={fieldListShown}
          className="flex items-center gap-1 rounded px-1.5 py-0.5 hover:bg-slate-100 hover:text-ink"
          title={fieldListShown ? "Hide the PivotTable Fields pane" : "Show the PivotTable Fields pane"}
        >
          {fieldListShown ? <PanelRightClose className="h-3.5 w-3.5" /> : <PanelRightOpen className="h-3.5 w-3.5" />}
          Field List
        </button>

        <div className="flex items-center gap-1" role="group" aria-label="Zoom">
          <button
            type="button"
            onClick={() => stepZoom(-1)}
            disabled={zoomIndex === 0}
            className="rounded p-0.5 hover:bg-slate-100 hover:text-ink disabled:opacity-40"
            aria-label="Zoom out"
            title="Zoom out"
          >
            <Minus className="h-3.5 w-3.5" />
          </button>
          <input
            type="range"
            min={0}
            max={ZOOM_LEVELS.length - 1}
            step={1}
            value={zoomIndex >= 0 ? zoomIndex : ZOOM_LEVELS.indexOf(1)}
            onChange={(e) => onZoomChange(ZOOM_LEVELS[Number(e.target.value)])}
            className="h-1 w-24 cursor-pointer accent-gold"
            aria-label="Zoom level"
            aria-valuetext={`${Math.round(zoom * 100)}%`}
          />
          <button
            type="button"
            onClick={() => stepZoom(1)}
            disabled={zoomIndex === ZOOM_LEVELS.length - 1}
            className="rounded p-0.5 hover:bg-slate-100 hover:text-ink disabled:opacity-40"
            aria-label="Zoom in"
            title="Zoom in"
          >
            <Plus className="h-3.5 w-3.5" />
          </button>
          {/* Click the percentage to go back to 100%, like Excel's Zoom to 100%. */}
          <button
            type="button"
            onClick={() => onZoomChange(1)}
            className="w-10 rounded py-0.5 text-right tabular-nums hover:bg-slate-100 hover:text-ink"
            title="Reset zoom to 100%"
          >
            {Math.round(zoom * 100)}%
          </button>
        </div>
      </div>
    </footer>
  );
}

function Stat({ label, value }) {
  return (
    <div className="flex items-center gap-1">
      <dt className="text-slate-500">{label}:</dt>
      <dd className="font-semibold text-ink">{value}</dd>
    </div>
  );
}
