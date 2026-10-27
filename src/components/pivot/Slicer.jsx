import { useEffect, useState } from "react";
import { FilterX, ListChecks } from "lucide-react";
import { getMembers } from "../../services/cubeMetaApi";
import { queryPivot } from "../../services/pivotApi";
import { itemsWithData } from "../../utils/slicerProbe";
import ContextMenu from "../forms/ContextMenu";
import { widgetTitle } from "../../utils/widgetTitle";

// Excel's own Slicer: a floating button list bound to one field, living OUTSIDE the Filters/Report-
// Filter-row area (Insert > Slicer). It reads and writes the SAME request.filters entry (Members
// mode) a Report Filter row would — just shown as buttons instead of a dropdown.
//
// Matched to Excel exactly (2026-10-25, the user: "like Excel — don't assume anything"): a header
// with the field name and only Multi-Select and Clear Filter; no search box; every item highlighted
// while the field is unfiltered; a click selects that item ALONE, Ctrl+click (or Multi-Select on)
// adds/removes items; right-click for Clear Filter / Remove, as Excel has no close button.
// probeRequest (utils/slicerProbe): the query telling which items have data under the other
// filters — those without are greyed out and listed last, like Excel's default slicer settings.
export default function Slicer({ field, displayName, filter, onChange, onRemove, probeRequest = null }) {
  const [members, setMembers] = useState([]);
  const [loading, setLoading] = useState(false);
  // A failed load (the shared cube's "busy" spells) used to show "No items" as if the field were
  // empty (2026-10-25): it gets its own message, a Retry, and an automatic retry after 15s.
  const [failed, setFailed] = useState(false);
  const [retryToken, setRetryToken] = useState(0);
  const [multiSelect, setMultiSelect] = useState(false);
  const [menu, setMenu] = useState(null);
  const includedMembers = filter?.includedMembers ?? [];
  const filtered = includedMembers.length > 0;
  const title = widgetTitle(displayName);

  useEffect(() => {
    let cancelled = false;
    const controller = new AbortController();
    setLoading(true);
    setFailed(false);
    let retryId = null;
    getMembers(field, undefined, controller.signal)
      .then((result) => {
        if (!cancelled) setMembers(result);
      })
      .catch((err) => {
        if (cancelled || err.name === "AbortError") return;
        setFailed(true);
        retryId = setTimeout(() => setRetryToken((t) => t + 1), 15000);
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
      clearTimeout(retryId);
      controller.abort();
    };
  }, [field, retryToken]);

  // Items with data under the other filters (null = not known, nothing greyed — also on a failed
  // query, so a busy server never makes every item look empty).
  const [withData, setWithData] = useState(null);
  const probeKey = probeRequest ? JSON.stringify(probeRequest) : "";
  useEffect(() => {
    if (!probeKey) {
      setWithData(null);
      return undefined;
    }
    const controller = new AbortController();
    const timer = setTimeout(() => {
      queryPivot(JSON.parse(probeKey), { signal: controller.signal })
        .then((result) => setWithData(itemsWithData(result)))
        .catch(() => setWithData(null));
    }, 400);
    return () => {
      clearTimeout(timer);
      controller.abort();
    };
  }, [probeKey]);
  const hasData = (member) => withData === null || withData.has(member);
  // Items with no data go last, keeping the field's order within each group.
  const orderedMembers = withData === null ? members : [...members.filter(hasData), ...members.filter((m) => !hasData(m))];

  const clearFilter = () => onChange({ field, mode: "Members", includedMembers: [] });

  // Excel: a plain click keeps only that item; Ctrl+click or Multi-Select toggles it. Toggling off
  // the last selected item, or selecting every item, is the same as no filter.
  const clickItem = (member, ctrlKey) => {
    let next;
    if (multiSelect || ctrlKey) {
      const current = filtered ? includedMembers : members;
      next = current.includes(member) ? current.filter((m) => m !== member) : [...current, member];
    } else {
      next = [member];
    }
    if (next.length === 0 || (members.length > 0 && members.every((m) => next.includes(m)))) clearFilter();
    else onChange({ field, mode: "Members", includedMembers: next });
  };

  const headerButton = "rounded p-1 text-muted hover:bg-slate-100 hover:text-ink disabled:cursor-not-allowed disabled:opacity-40";

  return (
    <div
      className="flex h-full w-full flex-col rounded-lg border border-slate-300 bg-white"
      onContextMenu={(e) => {
        e.preventDefault();
        setMenu({ x: e.clientX, y: e.clientY });
      }}
    >
      {/* The header is where the floating slicer is dragged from (PivotGrid's FloatingWidget). */}
      <div data-drag-handle className="flex cursor-move items-center justify-between gap-2 px-2.5 pb-1.5 pt-2">
        <span className="truncate text-sm font-semibold text-ink" title={displayName}>
          {title}
        </span>
        <div className="flex shrink-0 items-center gap-0.5">
          <button
            type="button"
            onClick={() => setMultiSelect((v) => !v)}
            aria-pressed={multiSelect}
            title="Multi-Select (Alt+S)"
            className={`${headerButton} ${multiSelect ? "bg-primary-soft text-primary-strong ring-1 ring-inset ring-primary/40" : ""}`}
          >
            <ListChecks className="h-4 w-4" />
          </button>
          <button type="button" onClick={clearFilter} disabled={!filtered} title="Clear Filter (Alt+C)" className={headerButton}>
            <FilterX className="h-4 w-4" />
          </button>
        </div>
      </div>

      <div
        className="custom-scrollbar flex min-h-0 flex-1 flex-col gap-1 overflow-y-auto px-2 pb-2"
        onKeyDown={(e) => {
          if (e.altKey && e.key.toLowerCase() === "s") setMultiSelect((v) => !v);
          if (e.altKey && e.key.toLowerCase() === "c") clearFilter();
        }}
      >
        {loading && members.length === 0 ? (
          <p className="py-2 text-center text-xs text-muted">Loading…</p>
        ) : failed && members.length === 0 ? (
          <div className="flex flex-col items-center gap-1.5 py-2 text-center text-xs text-rose-700">
            <span>Couldn't load the items — the data server is busy.</span>
            <button type="button" onClick={() => setRetryToken((t) => t + 1)} className="rounded border border-rose-300 px-2 py-0.5 font-medium hover:bg-rose-50">
              Retry
            </button>
          </div>
        ) : members.length === 0 ? (
          <p className="py-2 text-center text-xs text-muted">No items</p>
        ) : (
          orderedMembers.map((member) => {
            // Unfiltered, every item shows as selected — Excel's slicer look.
            const selected = !filtered || includedMembers.includes(member);
            const noData = !hasData(member);
            return (
              <button
                key={member}
                type="button"
                onClick={(e) => clickItem(member, e.ctrlKey || e.metaKey)}
                title={member === "" ? "(blank)" : member}
                aria-pressed={selected}
                // shrink-0: with more items than room the list scrolls, as in Excel, instead of
                // squashing every button until its text is cut off.
                className={`shrink-0 truncate rounded border px-2 py-1.5 text-left text-xs ${
                  noData
                    ? selected
                      ? "border-primary/30 bg-primary-50 text-slate-400"
                      : "border-dashed border-slate-200 bg-white text-slate-300"
                    : selected
                      ? "border-primary/60 bg-primary-100 text-ink"
                      : "border-slate-200 bg-white text-slate-400 hover:border-slate-300"
                }`}
              >
                {/* A member with an empty caption reads "(blank)", like Excel (display only). */}
                {member === "" ? "(blank)" : member}
              </button>
            );
          })
        )}
      </div>

      {menu && (
        <ContextMenu x={menu.x} y={menu.y} onClose={() => setMenu(null)}>
          <div className="flex flex-col text-sm">
            <button
              type="button"
              disabled={!filtered}
              onClick={() => {
                clearFilter();
                setMenu(null);
              }}
              className="rounded-lg px-2 py-1.5 text-left hover:bg-primary/10 disabled:cursor-not-allowed disabled:opacity-40"
            >
              Clear Filter From &quot;{title}&quot;
            </button>
            <button
              type="button"
              onClick={() => {
                setMenu(null);
                onRemove();
              }}
              className="rounded-lg px-2 py-1.5 text-left hover:bg-primary/10"
            >
              Remove &quot;{title}&quot;
            </button>
          </div>
        </ContextMenu>
      )}
    </div>
  );
}
