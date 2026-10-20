import { useEffect } from "react";
import { X } from "lucide-react";

// A centered, full-viewport modal — distinct from Popover.jsx (a small anchored panel next to its
// trigger): this is for content too big to float near a button, like a whole pivot table preview
// (see SnapshotHistory.jsx). Click-outside and Escape both close it, same as Popover.
export default function Modal({ title, onClose, children }) {
  useEffect(() => {
    const handleKeyDown = (e) => {
      if (e.key === "Escape") onClose();
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  }, [onClose]);

  return (
    // Full viewport since the left sidebar became a top bar (2026-10-20) — it used to start at the
    // sidebar edge (left-16/left-64) so it centred within the content area.
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-ink/60 p-4"
      onMouseDown={onClose}
    >
      <div
        // w-fit (not w-full): a narrow table left the card stretched to max-w-5xl regardless,
        // with a big empty gap to the right of the actual content — caught live 2026-09-23, same
        // "card should hug its content" fix PivotGrid's own wrapper already uses. Still caps at
        // max-w-5xl (falling back to internal scrolling, via PivotGrid's own overflow-x-auto) for
        // a table wide enough to need it.
        //
        // overflow-hidden here (not overflow-auto) — the actual scrolling moved to the inner div
        // just below. A scrollable element's native scrollbar doesn't respect its own
        // border-radius (it's drawn in the corner "notch" the radius creates), which squared off
        // this card's top-right/bottom-right corners wherever the scrollbar happened to sit
        // (caught live 2026-10-07, "عايز ال border radius من كل الجوانب"). This OUTER div now only
        // clips to the rounded shape; the INNER div does the actual scrolling, so its scrollbar
        // sits safely inside the already-rounded boundary instead of cutting into it.
        className="max-h-[90vh] w-fit max-w-5xl overflow-hidden rounded-lg bg-white shadow-2xl"
        onMouseDown={(e) => e.stopPropagation()}
      >
        {/* custom-scrollbar (see index.css) replaces the browser's own default scrollbar — also
            requested live 2026-10-07 alongside the rounded-corners fix above. */}
        <div className="custom-scrollbar max-h-[90vh] overflow-auto p-6">
          <div className="mb-4 flex items-center justify-between gap-4">
            <h3 className="text-lg font-bold text-ink">{title}</h3>
            <button type="button" onClick={onClose} className="rounded-full p-1 text-muted hover:bg-slate-100 hover:text-ink" aria-label="Close">
              <X className="h-5 w-5" />
            </button>
          </div>
          {children}
        </div>
      </div>
    </div>
  );
}
