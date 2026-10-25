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
//
// Real Excel's own filter checklist never re-queries per checkbox click — but that staging lives at
// the WHOLE-DIALOG level (an outer OK/Cancel covering field/mode/text/number inputs too, not just
// this checklist — see FilterRow's own WHY comment in FieldPicker.jsx), not inside this one widget.
// This component itself stays a plain, immediate controlled input (every click calls `onChange`
// straight away) — a brief Apply-button-only version of this lived here 2026-10-12 but was
// superseded the same day once the real Excel screenshot showed OK/Cancel belongs one level up, at
// the FilterRow/FilterPill popover itself, covering every control in it, not nested per-widget.
// `allowMultiple` (new, default false): real Excel's "Select Multiple Items" checkbox — unchecked,
// clicking any item REPLACES the whole selection with just that one (single-select), matching
// Excel's own default; checked, unlimited multi-check (today's only-ever-available behavior).
//
// `showSelectMultiple` (default true): confirmed live 2026-10-12 — this checkbox only exists on
// real Excel's REPORT FILTER (page field) dropdown. A ROW/COLUMN field's own filter dropdown has no
// such toggle at all — it's always unrestricted multi-select, since a Row/Column field's whole job
// is grouping multiple members. Callers filtering a Row/Column field's own members (FieldHeaderMenu,
// via FilterRow's `fieldLocked`) pass `showSelectMultiple={false}`, which also forces `allowMultiple`
// permanently true (no toggle to turn it on otherwise).
//
// `showAllItem` (2026-10-21): Excel's Report Filter list starts with "(All)" — the unfiltered state,
// an empty `value`. Picking it clears the selection; the caller treats an empty list as no filter.
export default function MemberCheckboxList({ options, value, onChange, placeholder, isLoading = false, onInputChange, showSelectMultiple = true, showAllItem = false, onAllowMultipleChange }) {
  const [search, setSearch] = useState("");
  // Opens ticked when the filter already holds several items (as Excel shows it) — unticked, the
  // next click would silently replace them all with one.
  const [allowMultipleState, setAllowMultiple] = useState(() => value.length > 1);
  const allowMultiple = showSelectMultiple ? allowMultipleState : true;
  const selected = new Set(value);
  // onInputChange present (FilterRow's own debounced getMembers re-fetch): `options` IS already the
  // server-filtered result, rendered as-is. Without it (Calculated Item's Add/Subtract, which fetch
  // the full member list ONCE up front — see their own call sites): `options` is the WHOLE list, so
  // this filters it client-side instead, the same way the tag-style MultiSelectField this replaces
  // always relied on react-select's own built-in client-side filtering for that exact case.
  const visibleOptions = onInputChange ? options : options.filter((o) => o.toLowerCase().includes(search.toLowerCase()));
  // "(Select All)" only ever acts on the CURRENTLY VISIBLE (searched/filtered) options — same as
  // real Excel's own filter checklist, not literally every member of the field regardless of search.
  // Only offered in multi-select mode — Excel has no "select all" concept in its single-select mode.
  const allVisibleSelected = visibleOptions.length > 0 && visibleOptions.every((o) => selected.has(o));

  const toggleOne = (option) => {
    if (!allowMultiple) {
      onChange(selected.has(option) ? [] : [option]);
      return;
    }
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
      <div className="flex items-center gap-1.5 rounded-xl border border-slate-200 bg-white px-3 py-2 focus-within:ring-2 focus-within:ring-primary">
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
            {showAllItem && !allowMultiple && !search && (
              <label className="flex cursor-pointer items-center gap-2 border-b border-slate-100 px-3 py-1.5 text-sm text-ink hover:bg-slate-50">
                <input type="radio" className="h-4 w-4 accent-primary rounded-full" checked={value.length === 0} onChange={() => onChange([])} />
                (All)
              </label>
            )}
            {allowMultiple && (
              <label className="flex cursor-pointer items-center gap-2 border-b border-slate-100 bg-slate-50 px-3 py-2 text-sm font-semibold text-ink">
                <input type="checkbox" className="h-4 w-4 accent-primary" checked={allVisibleSelected} onChange={toggleAll} />
                (Select All)
              </label>
            )}
            {visibleOptions.map((option) => (
              <label key={option} className="flex cursor-pointer items-center gap-2 px-3 py-1.5 text-sm text-ink hover:bg-slate-50">
                <input
                  type={allowMultiple ? "checkbox" : "radio"}
                  className={allowMultiple ? "h-4 w-4 accent-primary" : "h-4 w-4 accent-primary rounded-full"}
                  checked={selected.has(option)}
                  onChange={() => toggleOne(option)}
                />
                {/* Excel lists a member with an empty caption as "(blank)" (display only — the value stays ""). */}
                <span className="truncate">{option === "" ? "(blank)" : option}</span>
              </label>
            ))}
          </>
        )}
      </div>
      {/* Positioned AFTER the list, right before FilterRow's own OK/Cancel — matches real Excel's
          own layout exactly (search box, then the list, then this checkbox, then OK/Cancel), not
          between the search box and the list where an earlier version of this had it (caught live,
          2026-10-12, directly against the real Excel screenshot: "متاكد ان دا الشكل بالظبط"). A
          "N selected" count line used to sit here too — removed the same way, confirmed by that
          same screenshot to not exist in real Excel's dropdown at all (nothing between this
          checkbox and OK/Cancel). */}
      {showSelectMultiple && (
        <label className="flex cursor-pointer items-center gap-2 text-xs text-ink">
          <input
            type="checkbox"
            className="h-3.5 w-3.5 accent-primary"
            checked={allowMultipleState}
            onChange={(e) => {
              setAllowMultiple(e.target.checked);
              onAllowMultipleChange?.(e.target.checked);
              // From "(All)", Excel shows every item checked — so unticking one leaves the rest.
              if (e.target.checked && showAllItem && value.length === 0 && !search) onChange(options);
            }}
          />
          Select Multiple Items
        </label>
      )}
    </div>
  );
}
