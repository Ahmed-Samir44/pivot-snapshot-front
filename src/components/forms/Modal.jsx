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
    // left-64 (not inset-0's implicit left-0): AppShell's sidebar is a permanently-visible
    // fixed w-64 column, so centering against the FULL viewport (including the space the
    // sidebar covers) put the modal visibly left of center within the actual content area —
    // caught live 2026-09-23. Centering only across the content area (viewport minus the
    // sidebar) matches what looks centered on screen.
    <div className="fixed inset-y-0 left-64 right-0 z-50 flex items-center justify-center bg-ink/40 p-4" onMouseDown={onClose}>
      <div
        // w-fit (not w-full): a narrow table left the card stretched to max-w-5xl regardless,
        // with a big empty gap to the right of the actual content — caught live 2026-09-23, same
        // "card should hug its content" fix PivotGrid's own wrapper already uses. Still caps at
        // max-w-5xl (falling back to internal scrolling, via PivotGrid's own overflow-x-auto) for
        // a table wide enough to need it.
        className="max-h-[90vh] w-fit max-w-5xl overflow-auto rounded-xl bg-white p-6 shadow-2xl"
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className="mb-4 flex items-center justify-between gap-4">
          <h3 className="text-lg font-bold text-ink">{title}</h3>
          <button type="button" onClick={onClose} className="rounded-full p-1 text-muted hover:bg-slate-100 hover:text-ink" aria-label="Close">
            <X className="h-5 w-5" />
          </button>
        </div>
        {children}
      </div>
    </div>
  );
}
