import { useEffect, useRef, useState } from "react";

// A minimal anchored popover: click the trigger to open a floating panel below it, click
// anywhere outside (or call the passed `close`) to dismiss. Used to move a zone's "add/edit"
// form out of the always-visible compact drop-zone box and into an on-demand panel — see
// FieldPicker.jsx's compact Rows/Columns/Values/Filters boxes and its Sort/Totals/Calculated
// Fields/Calculated Items controls, all of which reuse this instead of their own popover logic.
export default function Popover({ trigger, children, align = "left", panelClassName = "", inline = false }) {
  const [open, setOpen] = useState(false);
  const containerRef = useRef(null);

  useEffect(() => {
    if (!open) return undefined;
    const handlePointerDown = (e) => {
      if (containerRef.current && !containerRef.current.contains(e.target)) {
        setOpen(false);
      }
    };
    const handleKeyDown = (e) => {
      if (e.key === "Escape") setOpen(false);
    };
    document.addEventListener("mousedown", handlePointerDown);
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("mousedown", handlePointerDown);
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [open]);

  const close = () => setOpen(false);

  return (
    <div className={inline ? "relative" : "relative inline-block"} ref={containerRef}>
      {trigger(() => setOpen((o) => !o), open)}
      {open && (
        <div
          // onMouseDown here, not just the document-level "outside click" listener above: an
          // absolutely-positioned panel (the non-inline case) has nothing reserving space for it,
          // so an open panel near the bottom of a section can visually cover whatever sits below it
          // in the page. A click on that covered area lands INSIDE this div's own padding —
          // containerRef.contains(e.target) is true, so the document listener doesn't treat it as
          // "outside" and the click was just silently absorbed, leaving the covered element looking
          // broken/unclickable (caught live, 2026-09-27, on the toolbar's Sort popover covering
          // "Run query"). Closing on a mousedown that lands on the panel's own background (not
          // bubbled up from an actual control inside it, per the e.target === e.currentTarget
          // check) reveals what's underneath on that same click, so the very next click reaches it.
          // Kept for the inline case too even though nothing is covered there — harmless, and still
          // closes the panel on a stray click into its own padding.
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) setOpen(false);
          }}
          className={
            inline
              ? // inline: a normal block below the trigger, IN FLOW — the toolbar row (Sort/
                // Totals/No-Data Items/Calculated Field/Calculated Item, all sitting directly above
                // the "Run query" button with barely any gap) needs opening one of these to actually
                // push Run query down, not float an absolutely-positioned panel over it regardless
                // of panel height (the onMouseDown workaround above only helps when the covered spot
                // is blank panel padding — it did nothing when Calculated Field's own "+ Add
                // calculated field" button was what covered Run query, since that's a real control,
                // not blank space — caught live immediately after the first fix, 2026-09-27).
                `mt-2 min-w-[22rem] max-w-[26rem] rounded-xl border border-slate-200 bg-white p-4 shadow-xl ${panelClassName}`
              : `absolute z-50 mt-2 ${align === "right" ? "right-0" : "left-0"} min-w-[22rem] max-w-[26rem] rounded-xl border border-slate-200 bg-white p-4 shadow-xl ${panelClassName}`
          }
        >
          {children(close)}
        </div>
      )}
    </div>
  );
}
