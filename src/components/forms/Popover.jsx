import { useEffect, useRef, useState } from "react";

// A minimal anchored popover: click the trigger to open a floating panel below it, click
// anywhere outside (or call the passed `close`) to dismiss. Used to move a zone's "add/edit"
// form out of the always-visible compact drop-zone box and into an on-demand panel — see
// FieldPicker.jsx's compact Rows/Columns/Values/Filters boxes and its Sort/Totals/Calculated
// Fields/Calculated Items controls, all of which reuse this instead of their own popover logic.
export default function Popover({ trigger, children, align = "left", panelClassName = "" }) {
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
    <div className="relative inline-block" ref={containerRef}>
      {trigger(() => setOpen((o) => !o), open)}
      {open && (
        <div
          className={`absolute z-50 mt-2 ${align === "right" ? "right-0" : "left-0"} min-w-[22rem] max-w-[26rem] rounded-xl border border-slate-200 bg-white p-4 shadow-xl ${panelClassName}`}
        >
          {children(close)}
        </div>
      )}
    </div>
  );
}
