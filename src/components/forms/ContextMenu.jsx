import { useEffect, useLayoutEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";

// A right-click-anchored menu — portaled to document.body with `position: fixed`, same escape-the-
// ancestor approach Popover.jsx now uses (see its own WHY comment), but anchored to a POINT (the
// triggering contextmenu event's clientX/clientY) instead of a trigger element's rect, which is why
// this is its own small component rather than an extra mode bolted onto Popover's trigger-rect API.
// Click-outside/Escape close it, same as Popover and Modal.
export default function ContextMenu({ x, y, onClose, children }) {
  const menuRef = useRef(null);
  const [position, setPosition] = useState({ top: y, left: x });

  useEffect(() => {
    const handlePointerDown = (e) => {
      if (menuRef.current && !menuRef.current.contains(e.target)) onClose();
    };
    const handleKeyDown = (e) => {
      if (e.key === "Escape") onClose();
    };
    document.addEventListener("mousedown", handlePointerDown);
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("mousedown", handlePointerDown);
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [onClose]);

  // Keeps the menu fully on-screen — a right-click near the bottom/right edge of the viewport would
  // otherwise open a menu that spills past it. Measured after the first paint (menuRef has real
  // dimensions by then), re-measured whenever the menu's own content changes size (e.g. swapping
  // from the root list to a submenu panel with different height — the Number Format panel is both
  // taller AND wider than the root Sort/Format Cells/.../Remove list it replaces).
  //
  // The original version here only ever SLID the panel up just far enough to keep its bottom edge
  // on-screen, still anchored with its TOP edge wherever the click happened to land — for a click
  // near the bottom of the table, a tall panel (Number Format's category list) got dragged far
  // enough up the screen to render well ABOVE the clicked cell, with the table's own rows showing
  // through/around it looking exactly like Popover.jsx's own now-fixed "panel doesn't react to its
  // own size changing" class of bug (caught live, 2026-10-06, as a repeat of that SAME complaint —
  // "بردو طالعه برا المنيو مع اني قولتلك قبل كدا" — this component was never actually touched by
  // that earlier fix, which only landed in Popover.jsx; this is a structurally separate component
  // since it anchors to a POINT, not a trigger element's rect, so it needed the same fix applied
  // here too). Switched to Popover.jsx's own flip-upward approach instead: open the panel ABOVE the
  // click point entirely (bottom edge at y, not top edge at y) whenever there's more room above than
  // below and not enough room below for the panel's real height — plus a genuine two-sided left/top
  // clamp (Popover.jsx's `Math.min(Math.max(..., 8), max)` pattern) instead of only ever correcting
  // the right/bottom edges.
  useLayoutEffect(() => {
    const el = menuRef.current;
    if (!el) return undefined;
    const reposition = () => {
      const rect = el.getBoundingClientRect();
      const panelWidth = rect.width;
      const panelHeight = rect.height;
      const maxLeft = Math.max(8, window.innerWidth - panelWidth - 8);
      const left = Math.min(Math.max(x, 8), maxLeft);

      const spaceBelow = window.innerHeight - y;
      const spaceAbove = y;
      const openUpward = panelHeight > 0 && spaceBelow < panelHeight + 8 && spaceAbove > spaceBelow;
      const maxTop = Math.max(8, window.innerHeight - panelHeight - 8);
      const top = openUpward ? Math.max(8, y - panelHeight) : Math.min(y, maxTop);
      setPosition({ top, left });
    };
    reposition();
    const observer = new ResizeObserver(reposition);
    observer.observe(el);
    return () => observer.disconnect();
  }, [x, y]);

  return createPortal(
    <div
      ref={menuRef}
      style={{ position: "fixed", top: position.top, left: position.left }}
      className="z-50 min-w-[14rem] max-w-[20rem] rounded-xl border border-slate-200 bg-white p-1 text-sm shadow-xl"
      onMouseDown={(e) => e.stopPropagation()}
    >
      {children}
    </div>,
    document.body,
  );
}
