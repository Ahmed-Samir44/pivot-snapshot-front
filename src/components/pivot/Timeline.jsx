import { useEffect, useMemo, useState } from "react";
import { X, FilterX } from "lucide-react";
import { getMembers } from "../../services/cubeMetaApi";

const GRANULARITIES = ["Years", "Quarters", "Months", "Days"];

// The cube's own date dimensions return one member per raw calendar date, captioned "M/d/yyyy"
// (confirmed live 2026-09-22, see MdxPivotQueryBuilder's own WHY comment on date field identifier
// shapes) — parsed here the same explicit way rather than new Date(caption)/Date.parse, both of
// which are locale/engine-dependent for an ambiguous "M/d/yyyy" string (some treat it as d/M).
function parseMemberDate(caption) {
  const match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(caption);
  if (!match) return null;
  const [, month, day, year] = match;
  return new Date(Number(year), Number(month) - 1, Number(day));
}

function nextPeriod(date, granularity) {
  const next = new Date(date);
  if (granularity === "Years") next.setFullYear(next.getFullYear() + 1);
  else if (granularity === "Quarters") next.setMonth(next.getMonth() + 3);
  else if (granularity === "Months") next.setMonth(next.getMonth() + 1);
  else next.setDate(next.getDate() + 1);
  return next;
}

function bucketKey(date, granularity) {
  if (granularity === "Years") return `${date.getFullYear()}`;
  if (granularity === "Quarters") return `${date.getFullYear()} Q${Math.floor(date.getMonth() / 3) + 1}`;
  if (granularity === "Months") return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function bucketLabel(key, granularity) {
  if (granularity === "Months") {
    const [year, month] = key.split("-");
    return new Date(Number(year), Number(month) - 1, 1).toLocaleString("en-US", { month: "short", year: "numeric" });
  }
  if (granularity === "Days") {
    const [year, month, day] = key.split("-");
    return new Date(Number(year), Number(month) - 1, Number(day)).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
  }
  return key;
}

// A bucket's first and last day, as ISO "yyyy-mm-dd" — what a DateBetween filter takes.
const iso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
function bucketRange(key, granularity) {
  if (granularity === "Years") return [new Date(Number(key), 0, 1), new Date(Number(key), 11, 31)];
  if (granularity === "Quarters") {
    const [year, q] = key.split(" Q").map(Number);
    return [new Date(year, (q - 1) * 3, 1), new Date(year, q * 3, 0)];
  }
  if (granularity === "Months") {
    const [year, month] = key.split("-").map(Number);
    return [new Date(year, month - 1, 1), new Date(year, month, 0)];
  }
  const [year, month, day] = key.split("-").map(Number);
  const d = new Date(year, month - 1, day);
  return [d, d];
}

// Excel's own Timeline control (Insert > Timeline) — a date-range picker bound to one field,
// distinct from a plain Slicer in that it groups members into Years/Quarters/Months/Days "boxes"
// you click (or shift-click to extend) to select a CONTIGUOUS range, instead of toggling individual
// members one at a time. Like Slicer.jsx, this is purely a different widget over the SAME
// request.filters entry (Members mode) — the range selection is resolved to an explicit member list
// client-side, so the backend needs no new filter concept for it.
export default function Timeline({ field, displayName, filter, onChange, onRemove }) {
  const [allMembers, setAllMembers] = useState([]);
  const [loading, setLoading] = useState(true);
  // Matches real Excel's own default the moment you add a date field — it auto-groups into a
  // Year/Quarter/Month/Day hierarchy starting at Month, never listing every raw day up front
  // (confirmed live 2026-10-06, after an unfiltered day-level member fetch crashed the cube — see
  // this component's own field-choice note in PivotToolbar.jsx's Insert Timeline tooltip).
  const [granularity, setGranularity] = useState("Months");
  const [rangeStart, setRangeStart] = useState(null);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    getMembers(field)
      .then((members) => {
        if (!cancelled) setAllMembers(members);
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [field]);

  const dated = useMemo(
    () => allMembers.map((caption) => ({ caption, date: parseMemberDate(caption) })).filter((m) => m.date !== null),
    [allMembers],
  );

  // The full span from the EARLIEST real date through TODAY, not just whichever periods happen to
  // have at least one member — requested live, 2026-10-06 ("date picker should be from the
  // beginning of date to the current date"). A period with no actual data still gets its own box
  // (clicking it just contributes zero members to the filter, same as real Excel's own Timeline
  // showing every period in range regardless of whether it has data), so you can navigate all the
  // way to "now" instead of the control stopping wherever the last real row happened to land.
  const buckets = useMemo(() => {
    if (dated.length === 0) return [];
    const minDate = dated.reduce((min, m) => (m.date < min ? m.date : min), dated[0].date);
    const today = new Date();
    const keys = [];
    let cursor = minDate;
    while (cursor <= today) {
      keys.push(bucketKey(cursor, granularity));
      cursor = nextPeriod(cursor, granularity);
    }
    return keys;
  }, [dated, granularity]);

  // Which buckets the CURRENT filter covers — lets the timeline re-derive its own highlighted range
  // after a page reload instead of forgetting the selection. A date range (what the timeline sets
  // since 2026-10-21) covers the buckets inside it; an older saved member list, its members' buckets.
  const includedMembers = filter?.includedMembers ?? [];
  const isRange = filter?.mode === "DateBetween" && Boolean(filter.labelText) && Boolean(filter.labelTextTo);
  const rangeFrom = isRange ? filter.labelText : null;
  const rangeTo = isRange ? filter.labelTextTo : null;
  const hasSelection = isRange || includedMembers.length > 0;
  const selectedBuckets = useMemo(() => {
    if (rangeFrom && rangeTo) {
      return new Set(buckets.filter((b) => {
        const [start, end] = bucketRange(b, granularity);
        return iso(start) >= rangeFrom && iso(end) <= rangeTo;
      }));
    }
    if (includedMembers.length === 0) return new Set();
    const included = new Set(includedMembers);
    const keys = new Set();
    for (const { caption, date } of dated) {
      if (included.has(caption)) keys.add(bucketKey(date, granularity));
    }
    return keys;
  }, [includedMembers, dated, granularity, buckets, rangeFrom, rangeTo]);

  const selectedIndexes = buckets.map((b, i) => (selectedBuckets.has(b) ? i : -1)).filter((i) => i >= 0);
  const rangeLo = selectedIndexes.length > 0 ? Math.min(...selectedIndexes) : null;
  const rangeHi = selectedIndexes.length > 0 ? Math.max(...selectedIndexes) : null;

  // A DATE RANGE from the first bucket's first day to the last one's last day, like Excel's timeline
  // (2026-10-21) — it used to be the list of member dates in those buckets, which missed any date
  // the member list didn't have (found by the feature test: 2026 came out 2,640.25 short).
  const applyRange = (loIndex, hiIndex) => {
    const lo = Math.min(loIndex, hiIndex);
    const hi = Math.max(loIndex, hiIndex);
    const [start] = bucketRange(buckets[lo], granularity);
    const [, end] = bucketRange(buckets[hi], granularity);
    onChange({ field, mode: "DateBetween", includedMembers: [], labelText: iso(start), labelTextTo: iso(end) });
  };

  const handleBucketClick = (index, shiftKey) => {
    if (shiftKey && rangeStart !== null) {
      applyRange(rangeStart, index);
    } else {
      setRangeStart(index);
      applyRange(index, index);
    }
  };

  const clearFilter = () => {
    setRangeStart(null);
    onChange({ field, mode: "Members", includedMembers: [] });
  };

  return (
    <div className="flex w-full max-w-2xl flex-col rounded-xl border border-slate-300/70 bg-white shadow-sm">
      <div className="flex items-center justify-between gap-2 border-b border-slate-200 bg-slate-100 px-3 py-2">
        <span className="truncate text-xs font-bold text-ink" title={displayName}>
          {displayName}
        </span>
        <div className="flex shrink-0 items-center gap-2">
          <select
            value={granularity}
            onChange={(e) => {
              setRangeStart(null);
              setGranularity(e.target.value);
            }}
            className="rounded border border-slate-200 bg-white px-1.5 py-0.5 text-xs"
          >
            {GRANULARITIES.map((g) => (
              <option key={g} value={g}>
                {g}
              </option>
            ))}
          </select>
          <button
            type="button"
            onClick={clearFilter}
            disabled={!hasSelection}
            title="Clear Filter"
            className="rounded p-0.5 text-muted hover:bg-slate-200 hover:text-ink disabled:cursor-not-allowed disabled:opacity-40"
          >
            <FilterX className="h-3.5 w-3.5" />
          </button>
          <button type="button" onClick={onRemove} title="Remove Timeline" className="rounded p-0.5 text-muted hover:bg-slate-200 hover:text-ink">
            <X className="h-3.5 w-3.5" />
          </button>
        </div>
      </div>

      <div className="flex gap-1 overflow-x-auto p-2">
        {loading ? (
          <p className="py-2 text-xs text-muted">Loading…</p>
        ) : buckets.length === 0 ? (
          <p className="py-2 text-xs text-muted">This field has no date-shaped members.</p>
        ) : (
          buckets.map((bucket, index) => {
            const isSelected = rangeLo !== null && index >= rangeLo && index <= rangeHi;
            return (
              <button
                key={bucket}
                type="button"
                onClick={(e) => handleBucketClick(index, e.shiftKey)}
                title="Click to select, Shift+Click to extend the range"
                className={`min-w-[3.5rem] shrink-0 rounded px-2 py-3 text-center text-xs font-medium ${
                  isSelected ? "bg-gold text-white" : "bg-slate-50 text-ink hover:bg-slate-100"
                }`}
              >
                {bucketLabel(bucket, granularity)}
              </button>
            );
          })
        )}
      </div>
    </div>
  );
}
