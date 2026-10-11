import { useEffect, useState } from "react";
import { X } from "lucide-react";
import { readCollapsed, SIDEBAR_COLLAPSED_EVENT } from "../layout/AppShell";

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

  // Starts from whatever's persisted, then stays LIVE for the rest of this Modal's lifetime via
  // AppShell's own toggle event — a plain read-once-at-mount was still wrong for a Modal opened
  // BEFORE the sidebar was toggled (its backdrop offset would freeze at whatever the sidebar's
  // width was at mount time, not whatever it becomes afterward) — caught live 2026-10-07, "لسه
  // بايظ" against exactly that sequence.
  const [sidebarCollapsed, setSidebarCollapsed] = useState(readCollapsed);
  useEffect(() => {
    const handleChange = (e) => setSidebarCollapsed(e.detail);
    window.addEventListener(SIDEBAR_COLLAPSED_EVENT, handleChange);
    return () => window.removeEventListener(SIDEBAR_COLLAPSED_EVENT, handleChange);
  }, []);

  return (
    // left-16/left-64 (not inset-0's implicit left-0): AppShell's sidebar is a permanently-visible
    // fixed column whose width toggles between w-16 (collapsed) and w-64 (expanded) — centering
    // against the FULL viewport (including the space the sidebar covers) put the modal visibly off-
    // center within the actual content area (caught live 2026-09-23 for the always-expanded case;
    // the collapsed case was never accounted for at all until caught live 2026-10-07). bg-ink/70 (up
    // from /40): AppShell's own page background is a diagonal gradient (from-slate-50 via-blue-50
    // to-slate-100, see AppShell.jsx), which a LIGHTER translucent overlay let show through
    // unevenly across the backdrop's width — visible as a stray lighter band near the sidebar edge,
    // independent of (and still present after) the left-offset fix above. A darker, closer-to-
    // opaque overlay makes that underlying gradient imperceptible regardless of position, instead of
    // chasing the gradient's own exact color stops to try to compensate for it.
    <div
      className={`fixed inset-y-0 ${sidebarCollapsed ? "left-16" : "left-64"} right-0 z-50 flex items-center justify-center bg-ink/70 p-4`}
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
        className="max-h-[90vh] w-fit max-w-5xl overflow-hidden rounded-xl bg-white shadow-2xl"
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
