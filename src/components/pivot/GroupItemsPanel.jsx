import { useEffect, useMemo, useState } from "react";
import { ChevronLeft, Loader2 } from "lucide-react";
import { getMembers } from "../../services/cubeMetaApi";

// Excel's "Group Selection" / "Ungroup" for one Rows/Columns field (2026-10-20) — request.itemGroups:
// { field, name, members }. The backend makes each group a single item aggregated BY THE CUBE (right
// for distinct counts too) in place of its members. Like Excel: new groups are named Group1, Group2…
// and renamed in place; Ungroup puts the members back.
export default function GroupItemsPanel({ field, request, onChange, onBack, onDone }) {
  const groups = (request.itemGroups ?? []).filter((g) => g.field === field);
  const otherGroups = (request.itemGroups ?? []).filter((g) => g.field !== field);
  const grouped = useMemo(() => new Set(groups.flatMap((g) => g.members)), [groups]);
  const blockedBy =
    request.dateGroupings?.some((g) => g.field === field) || request.numericGroupings?.some((g) => g.field === field)
      ? "This field is grouped by date or number ranges — remove that grouping first."
      : null;

  const [members, setMembers] = useState(null);
  const [loadError, setLoadError] = useState(null);
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState(new Set());
  const nextName = () => {
    const taken = new Set([...groups.map((g) => g.name), ...(request.calculatedItems ?? []).filter((i) => i.field === field).map((i) => i.name)]);
    let n = 1;
    while (taken.has(`Group${n}`)) n++;
    return `Group${n}`;
  };
  const [name, setName] = useState(nextName);
  // Renames are edited locally and applied on blur/Enter, so typing doesn't re-run the query per key.
  const [renames, setRenames] = useState({});

  useEffect(() => {
    const controller = new AbortController();
    getMembers(field, undefined, controller.signal)
      .then((list) => setMembers(list))
      .catch((err) => {
        if (err.name !== "AbortError") setLoadError(err.message);
      });
    return () => controller.abort();
  }, [field]);

  const save = (nextGroups) => onChange({ ...request, itemGroups: [...otherGroups, ...nextGroups] });

  const available = (members ?? []).filter((m) => m !== "" && !grouped.has(m));
  const term = search.trim().toLowerCase();
  const shown = term ? available.filter((m) => m.toLowerCase().includes(term)) : available;
  const trimmedName = name.trim();
  const nameProblem = !trimmedName
    ? "Enter a name."
    : /[[\]]/.test(trimmedName)
      ? "A name can't contain [ or ]."
      : groups.some((g) => g.name === trimmedName) ||
          available.includes(trimmedName) ||
          grouped.has(trimmedName) ||
          (request.calculatedItems ?? []).some((i) => i.field === field && i.name === trimmedName)
        ? "That name is already an item of this field."
        : null;

  const toggle = (member) =>
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(member)) next.delete(member);
      else next.add(member);
      return next;
    });

  const createGroup = () => {
    save([...groups, { field, name: trimmedName, members: available.filter((m) => selected.has(m)) }]);
    setSelected(new Set());
    setName(nextName());
    onDone?.();
  };

  const commitRename = (group) => {
    const value = (renames[group.name] ?? group.name).trim();
    setRenames((current) => {
      const rest = { ...current };
      delete rest[group.name];
      return rest;
    });
    if (!value || value === group.name || /[[\]]/.test(value) || groups.some((g) => g.name === value)) return;
    save(groups.map((g) => (g === group ? { ...g, name: value } : g)));
  };

  return (
    <div className="space-y-3 text-sm">
      <button type="button" onClick={onBack} className="inline-flex items-center gap-1 text-xs font-medium text-muted hover:text-ink">
        <ChevronLeft className="h-3.5 w-3.5" /> Back
      </button>

      {blockedBy ? (
        <p className="text-xs text-muted">{blockedBy}</p>
      ) : (
        <>
          {groups.length > 0 && (
            <div className="space-y-1.5">
              <p className="text-xs font-semibold text-ink">Groups</p>
              <ul className="space-y-1">
                {groups.map((g) => (
                  <li key={g.name} className="flex items-center gap-2">
                    <input
                      className="input-field min-w-0 flex-1 py-1 text-xs"
                      value={renames[g.name] ?? g.name}
                      onChange={(e) => setRenames((r) => ({ ...r, [g.name]: e.target.value }))}
                      onBlur={() => commitRename(g)}
                      onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()}
                      aria-label={`Rename ${g.name}`}
                      title={g.members.join(", ")}
                    />
                    <span className="shrink-0 text-xs text-muted">{g.members.length} items</span>
                    <button
                      type="button"
                      onClick={() => save(groups.filter((x) => x !== g))}
                      className="shrink-0 text-xs font-semibold text-orange-600 hover:underline"
                    >
                      Ungroup
                    </button>
                  </li>
                ))}
              </ul>
            </div>
          )}

          <div className="space-y-1.5 border-t border-slate-100 pt-2">
            <p className="text-xs font-semibold text-ink">New group</p>
            {loadError ? (
              <p className="text-xs text-red-700">{loadError}</p>
            ) : members == null ? (
              <p className="flex items-center gap-1.5 text-xs text-muted">
                <Loader2 className="h-3.5 w-3.5 animate-spin" /> Loading items…
              </p>
            ) : (
              <>
                <input
                  type="search"
                  className="input-field w-full py-1 text-xs"
                  placeholder="Search items…"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                />
                <ul className="custom-scrollbar max-h-48 overflow-auto rounded-lg border border-slate-200 p-1">
                  {shown.map((m) => (
                    <li key={m}>
                      <label className="flex cursor-pointer items-center gap-2 rounded px-1.5 py-0.5 text-xs hover:bg-slate-50">
                        <input type="checkbox" className="h-3.5 w-3.5 accent-primary" checked={selected.has(m)} onChange={() => toggle(m)} />
                        {m}
                      </label>
                    </li>
                  ))}
                  {shown.length === 0 && <li className="px-1.5 py-1 text-xs text-muted">No items.</li>}
                </ul>
                <label className="flex items-center gap-2 text-xs">
                  <span className="shrink-0 font-medium text-muted">Name:</span>
                  <input className="input-field min-w-0 flex-1 py-1 text-xs" value={name} onChange={(e) => setName(e.target.value)} />
                </label>
                {selected.size > 0 && nameProblem && <p className="text-xs text-red-700">{nameProblem}</p>}
                <button
                  type="button"
                  onClick={createGroup}
                  disabled={selected.size === 0 || Boolean(nameProblem)}
                  className="btn-primary w-full py-1.5 text-xs disabled:cursor-not-allowed disabled:opacity-50"
                >
                  Group {selected.size > 0 ? `${selected.size} item${selected.size === 1 ? "" : "s"}` : "selected items"}
                </button>
              </>
            )}
          </div>
        </>
      )}
    </div>
  );
}
