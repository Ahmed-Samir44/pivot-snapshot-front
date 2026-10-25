import { useRef } from "react";

// A vertical drag handle between the worksheet and a pane docked on its RIGHT (the PivotTable
// Fields pane, 2026-10-20): dragging left widens the pane. Pointer-captured so a fast drag over the
// grid doesn't lose it; arrow keys nudge it, double-click restores the default width.
export default function PaneSplitter({ width, min, max, defaultWidth, onChange, label = "Resize pane" }) {
  const drag = useRef(null);
  const clamp = (value) => Math.min(max, Math.max(min, Math.round(value)));

  return (
    <div
      role="separator"
      aria-orientation="vertical"
      aria-label={label}
      aria-valuenow={width}
      aria-valuemin={min}
      aria-valuemax={max}
      tabIndex={0}
      title="Drag to resize · double-click to reset"
      className="group relative hidden w-3 shrink-0 cursor-col-resize touch-none select-none outline-none lg:block"
      onPointerDown={(e) => {
        e.preventDefault();
        e.currentTarget.setPointerCapture(e.pointerId);
        drag.current = { startX: e.clientX, startWidth: width };
      }}
      onPointerMove={(e) => {
        if (!drag.current) return;
        onChange(clamp(drag.current.startWidth - (e.clientX - drag.current.startX)));
      }}
      onPointerUp={() => {
        drag.current = null;
      }}
      onPointerCancel={() => {
        drag.current = null;
      }}
      onDoubleClick={() => onChange(defaultWidth)}
      onKeyDown={(e) => {
        if (e.key === "ArrowLeft") onChange(clamp(width + 16));
        else if (e.key === "ArrowRight") onChange(clamp(width - 16));
        else return;
        e.preventDefault();
      }}
    >
      {/* The visible line: a hairline that turns gold on hover/focus/drag. */}
      <span className="absolute inset-y-0 left-1/2 w-px -translate-x-1/2 bg-slate-200 transition-colors group-hover:w-[3px] group-hover:bg-primary group-focus-visible:w-[3px] group-focus-visible:bg-primary group-active:bg-primary" />
    </div>
  );
}
