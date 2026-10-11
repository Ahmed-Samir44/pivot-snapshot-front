import { useState } from "react";
import { Search } from "lucide-react";

// Real Excel's own PivotTable field filter is a checkbox list by default — a search box, a
// "(Select All)" checkbox, then one checkbox per member — NOT a tag-style multi-select dropdown.
// This replaces MultiSelectField.jsx (a react-select port from the Segmentation project) for every
// "pick a set of members" use (Filter's Members/Exclude mode, Calculated Item's Add/Subtract) to
// actually match that real UI (flagged live, 2026-10-06, after a screenshot comparison — "المنيوز
// multi select with check box"). Same prop contract as MultiSelectField (options/value/onChange/
// placeholder/isLoading/onInputChange) so it's a drop-in swap at every call site.
//
// Deliberately NOT used for Custom Order (FieldPicker.jsx's own click-to-sequence picker) — that
// picker's whole point is the CLICK ORDER itself, which a checkbox list has no way to convey
// (checking two boxes doesn't tell you which one was checked first); it keeps the tag-style
// MultiSelectField, where each tag's own left-to-right order already shows the sequence.
export default function MemberCheckboxList({ options, value, onChange, placeholder, isLoading = false, onInputChange }) {
  const [search, setSearch] = useState("");
  const selected = new Set(value);
  // onInputChange present (FilterRow's own debounced getMembers re-fetch): `options` IS already the
  // server-filtered result, rendered as-is. Without it (Calculated Item's Add/Subtract, which fetch
  // the full member list ONCE up front — see their own call sites): `options` is the WHOLE list, so
  // this filters it client-side instead, the same way the tag-style MultiSelectField this replaces
  // always relied on react-select's own built-in client-side filtering for that exact case.
  const visibleOptions = onInputChange ? options : options.filter((o) => o.toLowerCase().includes(search.toLowerCase()));
  // "(Select All)" only ever acts on the CURRENTLY VISIBLE (searched/filtered) options — same as
  // real Excel's own filter checklist, not literally every member of the field regardless of search.
  const allVisibleSelected = visibleOptions.length > 0 && visibleOptions.every((o) => selected.has(o));

  const toggleOne = (option) => {
    onChange(selected.has(option) ? value.filter((v) => v !== option) : [...value, option]);
  };

  const toggleAll = () => {
    if (allVisibleSelected) {
      const visible = new Set(visibleOptions);
      onChange(value.filter((v) => !visible.has(v)));
    } else {
      onChange([...new Set([...value, ...visibleOptions])]);
    }
  };

  const handleSearchChange = (e) => {
    const text = e.target.value;
    setSearch(text);
    onInputChange?.(text, { action: "input-change" });
  };

  return (
    <div className="space-y-2">
      <div className="flex items-center gap-1.5 rounded-xl border border-slate-200 bg-white px-3 py-2 focus-within:ring-2 focus-within:ring-gold">
        <Search className="h-4 w-4 shrink-0 text-muted" />
        <input type="text" value={search} onChange={handleSearchChange} placeholder={placeholder} className="w-full text-sm outline-none" />
      </div>
      <div className="max-h-56 overflow-y-auto rounded-xl border border-slate-200">
        {isLoading ? (
          <p className="px-3 py-2 text-sm text-muted">Loading…</p>
        ) : visibleOptions.length === 0 ? (
          <p className="px-3 py-2 text-sm text-muted">No items</p>
        ) : (
          <>
            <label className="flex cursor-pointer items-center gap-2 border-b border-slate-100 bg-slate-50 px-3 py-2 text-sm font-semibold text-ink">
              <input type="checkbox" className="h-4 w-4 accent-gold" checked={allVisibleSelected} onChange={toggleAll} />
              (Select All)
            </label>
            {visibleOptions.map((option) => (
              <label key={option} className="flex cursor-pointer items-center gap-2 px-3 py-1.5 text-sm text-ink hover:bg-slate-50">
                <input type="checkbox" className="h-4 w-4 accent-gold" checked={selected.has(option)} onChange={() => toggleOne(option)} />
                <span className="truncate">{option}</span>
              </label>
            ))}
          </>
        )}
      </div>
      {value.length > 0 && <p className="text-xs text-muted">{value.length} selected</p>}
    </div>
  );
}
