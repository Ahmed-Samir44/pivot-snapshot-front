import { useEffect, useState } from "react";
import { X, Search, FilterX } from "lucide-react";
import { getMembers } from "../../services/cubeMetaApi";
import { widgetTitle } from "./Timeline";

// Excel's own Slicer: a floating button-grid control bound to one field — every member is its own
// clickable toggle button (selected = highlighted), with a small search box and a "clear filter"
// action, living OUTSIDE the Filters/Report-Filter-row area entirely (Insert > Slicer, not part of
// the Filters drop zone). This app already has everything a Slicer needs to actually FILTER with —
// a Slicer is just a different widget bound to the SAME request.filters entry (Members mode) the
// Report Filter row/sidebar Filters zone already read and write — so this renders a filter's
// included-members state as buttons instead of a dropdown, nothing new on the data side.
export default function Slicer({ field, displayName, filter, onChange, onRemove }) {
  const [members, setMembers] = useState([]);
  const [loading, setLoading] = useState(false);
  // A failed load (the shared cube's "busy" spells) used to show "No items" as if the field were
  // empty (2026-10-25): it gets its own message, a Retry, and an automatic retry after 15s.
  const [failed, setFailed] = useState(false);
  const [retryToken, setRetryToken] = useState(0);
  const [search, setSearch] = useState("");
  const includedMembers = filter?.includedMembers ?? [];

  useEffect(() => {
    let cancelled = false;
    const controller = new AbortController();
    setLoading(true);
    setFailed(false);
    let retryId = null;
    const timeoutId = setTimeout(() => {
      getMembers(field, search || undefined, controller.signal)
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
    }, 300);
    return () => {
      cancelled = true;
      clearTimeout(timeoutId);
      clearTimeout(retryId);
      controller.abort();
    };
  }, [field, search, retryToken]);

  const toggleMember = (member) => {
    const next = includedMembers.includes(member) ? includedMembers.filter((m) => m !== member) : [...includedMembers, member];
    onChange({ field, mode: "Members", includedMembers: next });
  };

  const clearFilter = () => onChange({ field, mode: "Members", includedMembers: [] });

  return (
    <div className="flex w-56 shrink-0 flex-col rounded-xl border border-slate-200 bg-white shadow-sm">
      <div className="flex items-center justify-between gap-2 border-b border-slate-200 px-3 py-2">
        <span className="truncate text-xs font-bold text-ink" title={displayName}>
          {widgetTitle(displayName)}
        </span>
        <div className="flex shrink-0 items-center gap-1">
          <button
            type="button"
            onClick={clearFilter}
            disabled={includedMembers.length === 0}
            title="Clear Filter"
            className="rounded p-0.5 text-muted hover:bg-slate-200 hover:text-ink disabled:cursor-not-allowed disabled:opacity-40"
          >
            <FilterX className="h-3.5 w-3.5" />
          </button>
          <button type="button" onClick={onRemove} title="Remove Slicer" className="rounded p-0.5 text-muted hover:bg-slate-200 hover:text-ink">
            <X className="h-3.5 w-3.5" />
          </button>
        </div>
      </div>

      <div className="border-b border-slate-200 p-1.5">
        <div className="flex items-center gap-1.5 rounded-lg border border-slate-200 bg-slate-50 px-2 py-1">
          <Search className="h-3.5 w-3.5 shrink-0 text-muted" />
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search…"
            className="w-full bg-transparent text-xs outline-none"
          />
        </div>
      </div>

      <div className="custom-scrollbar flex max-h-48 flex-col gap-1 overflow-y-auto p-1.5">
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
          members.map((member) => {
            const selected = includedMembers.includes(member);
            return (
              <button
                key={member}
                type="button"
                onClick={() => toggleMember(member)}
                title={member}
                aria-pressed={selected}
                className={`truncate rounded-md border px-2 py-1.5 text-left text-xs ${
                  selected ? "border-primary bg-primary font-semibold text-white" : "border-slate-200 bg-white text-ink hover:border-primary/40 hover:bg-primary-50"
                }`}
              >
                {/* A member with an empty caption reads "(blank)", like Excel (display only). */}
                {member === "" ? "(blank)" : member}
              </button>
            );
          })
        )}
      </div>
    </div>
  );
}
