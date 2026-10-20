import { useEffect, useLayoutEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";

// A minimal anchored popover: click the trigger to open a floating panel below it, click
// anywhere outside (or call the passed `close`) to dismiss. Used to move a zone's "add/edit"
// form out of the always-visible compact drop-zone box and into an on-demand panel — see
// FieldPicker.jsx's compact Rows/Columns/Values/Filters boxes and its Sort/Totals/Calculated
// Fields/Calculated Items controls, all of which reuse this instead of their own popover logic.
//
// The non-inline panel is portaled to document.body and positioned with `position: fixed` off
// the trigger's own getBoundingClientRect — same escape-the-ancestor trick this codebase already
// uses for every react-select menu (menuPortalTarget={document.body}/menuPosition="fixed" in
// MultiSelectField.jsx and FieldPicker.jsx's field Select). A plain CSS `absolute` panel is a
// DESCENDANT of whatever ancestor it's nested in — PivotGrid.jsx wraps both the Report Filter row
// and the table itself in a `.card` (glass-panel: translucent bg-white/80 + backdrop-blur-xl) with
// `overflow-hidden`, and the table's own scroll area adds `overflow-auto` on top of that. An
// absolute panel nested that deep gets clipped by those ancestors' overflow boxes, with whatever
// survives the clip rendering through the translucent/blurred glass-panel background instead of a
// solid one — caught live, 2026-10-05, as FieldHeaderMenu's and FilterHeaderRow's popovers opening
// as a broken, half-transparent box bleeding into the table rows beneath them. Porting straight to
// body sidesteps every ancestor's overflow/backdrop-filter entirely, regardless of which card ends
// up hosting the trigger next.
export default function Popover({
  trigger,
  children,
  align = "left",
  panelClassName = "",
  inline = false,
  // The wrapper is "inline-block" (shrink-to-fit) by default, which conflicts with a trigger
  // that's meant to be "w-full" of ITS OWN parent (e.g. a field-list pill styled as a full-width
  // Excel-style button): a 100%-width child inside a shrink-to-fit inline-block has nothing real
  // to size itself against, so both the trigger and the absolutely-positioned panel anchored to
  // this wrapper end up mis-sized/mis-positioned (caught live, 2026-10-05: a Values/Rows/Columns
  // pill's own dropdown menu rendering cut off and spilling outside its card). Pass true whenever
  // the trigger itself is "w-full".
  fullWidth = false,
  // Controlled mode: pass the open state and a real useState setter (it's called with an updater
  // function too, same as the internal setState below) when something OUTSIDE this component
  // needs to open/close the panel programmatically — e.g. a pill's right-click "Field Settings…"
  // menu item opening the same Popover a click on the pill itself would. Omit both for the
  // original uncontrolled behavior every other caller already relies on.
  open: controlledOpen,
  onOpenChange,
}) {
  const [internalOpen, setInternalOpen] = useState(false);
  const isControlled = controlledOpen !== undefined;
  const open = isControlled ? controlledOpen : internalOpen;
  const setOpen = isControlled ? onOpenChange : setInternalOpen;
  const containerRef = useRef(null);
  const panelRef = useRef(null);
  // Fixed-position coordinates for the portaled (non-inline) panel, recomputed off the trigger's
  // own rect — null until the first measurement lands, so the portal doesn't flash at (0, 0).
  const [fixedPos, setFixedPos] = useState(null);

  useEffect(() => {
    if (!open) return undefined;
    const handlePointerDown = (e) => {
      const insideTrigger = containerRef.current && containerRef.current.contains(e.target);
      const insidePanel = panelRef.current && panelRef.current.contains(e.target);
      if (!insideTrigger && !insidePanel) setOpen(false);
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
  }, [open, setOpen]);

  // Keeps the portaled panel anchored to the trigger as the page (or any scrollable ancestor,
  // e.g. the pivot table's own internal scroll area) moves underneath it — `capture: true` on
  // scroll is required since scroll events don't bubble, but a capture-phase listener on the
  // document still sees every descendant's scroll on its way down to the actual target.
  //
  // `left` is always resolved to a real pixel value here (not a separate CSS `right` offset) so it
  // can be clamped against the PANEL's own rendered width — align="right" anchors the panel's right
  // edge to the trigger's right edge, which spills off the LEFT side of the viewport whenever that
  // trigger sits near the screen's left edge with a panel wider than the gap to its left (caught
  // live, 2026-10-05: a row/column field header near the page's left edge opened its Sort/Filter
  // menu cut off, every line's leading text sheared away by the browser window's own edge — not an
  // ancestor overflow/clipping bug like the one this file's own portal already fixed, just nothing
  // here was keeping the panel on-screen once it had a real width to overflow with).
  useLayoutEffect(() => {
    if (inline || !open) return undefined;
    const reposition = () => {
      if (!containerRef.current) return;
      const rect = containerRef.current.getBoundingClientRect();
      // 0 on the very first pass (the portal hasn't rendered the panel at any position yet, so
      // panelRef is still null) — the rAF below re-runs this once it has, correcting the guess
      // below to the panel's real size instead of treating it as 0×0.
      const panelRect = panelRef.current?.getBoundingClientRect();
      const panelWidth = panelRect?.width ?? 0;
      const panelHeight = panelRect?.height ?? 0;
      const desiredLeft = align === "right" ? rect.right - panelWidth : rect.left;
      const maxLeft = Math.max(8, window.innerWidth - panelWidth - 8);
      const left = Math.min(Math.max(desiredLeft, 8), maxLeft);
      // Opening straight down from the trigger (the only thing this ever did before) has nothing
      // stopping a TALL panel from running off the bottom of the viewport or straight over whatever
      // page content sits below the trigger — only ever noticeable once a panel got tall enough to
      // reach that content, e.g. NumberFormatEditor's Category list, caught live 2026-10-06
      // overlapping the Save Snapshot bar beneath the table. Flips to open ABOVE the trigger instead
      // whenever there's more room up there than down, same "flip" behavior any real dropdown/
      // popover library has; only actually used once panelHeight is known (second pass, same timing
      // as the width-dependent `left` above) — the first pass still opens downward as a reasonable
      // guess before that.
      const spaceBelow = window.innerHeight - rect.bottom;
      const spaceAbove = rect.top;
      const openUpward = panelHeight > 0 && spaceBelow < panelHeight + 8 && spaceAbove > spaceBelow;
      const top = openUpward ? Math.max(8, rect.top - panelHeight) : rect.bottom;
      setFixedPos({ top, left });
    };
    reposition();
    // Re-clamps again once the panel has a real width (see reposition's own comment) AND attaches a
    // ResizeObserver so switching to WIDER content mid-interaction re-clamps too — e.g. NumberFormat
    // Editor's Currency category revealing a wider Symbol dropdown than whatever was selected when
    // this popover first positioned itself, growing the panel past the edge it was clamped against
    // for its ORIGINAL narrower width (caught live, 2026-10-06 — the exact same class of "nothing
    // here reacts to the panel's own size changing" bug this file already fixed once for the
    // trigger's own rect, just never extended to the panel's content changing size too).
    let observer;
    const raf = requestAnimationFrame(() => {
      reposition();
      if (panelRef.current) {
        observer = new ResizeObserver(reposition);
        observer.observe(panelRef.current);
      }
    });
    window.addEventListener("resize", reposition);
    document.addEventListener("scroll", reposition, true);
    return () => {
      cancelAnimationFrame(raf);
      observer?.disconnect();
      window.removeEventListener("resize", reposition);
      document.removeEventListener("scroll", reposition, true);
    };
  }, [inline, open, align]);

  const close = () => setOpen(false);

  const panelClasses = inline
    ? // inline: a normal block below the trigger, IN FLOW — the toolbar row (Sort/
      // Totals/No-Data Items/Calculated Field/Calculated Item, all sitting directly above
      // the "Run query" button with barely any gap) needs opening one of these to actually
      // push Run query down, not float a panel over it regardless of panel height.
      `mt-2 min-w-[22rem] max-w-[26rem] rounded-lg border border-slate-200 bg-white p-4 shadow-lg ${panelClassName}`
    : `fixed z-50 mt-2 min-w-[22rem] max-w-[26rem] rounded-lg border border-slate-200 bg-white p-4 shadow-lg ${panelClassName}`;

  const panel = open && (
    <div
      ref={panelRef}
      style={inline || !fixedPos ? undefined : { top: fixedPos.top, left: fixedPos.left }}
      // onMouseDown here, not just the document-level "outside click" listener above: a floating
      // panel has nothing reserving space for it, so an open panel near the bottom of a section
      // can visually cover whatever sits below it in the page. Closing on a mousedown that lands
      // on the panel's own background (not bubbled up from an actual control inside it, per the
      // e.target === e.currentTarget check) reveals what's underneath on that same click, so the
      // very next click reaches it (caught live, 2026-09-27, on the toolbar's Sort popover
      // covering "Run query"). Kept for the inline case too — harmless, still closes on a stray
      // click into its own padding.
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) setOpen(false);
      }}
      className={panelClasses}
    >
      {children(close)}
    </div>
  );

  return (
    <div className={inline ? "relative" : `relative ${fullWidth ? "block w-full" : "inline-block"}`} ref={containerRef}>
      {trigger(() => setOpen((o) => !o), open)}
      {inline ? panel : open && fixedPos && createPortal(panel, document.body)}
    </div>
  );
}
