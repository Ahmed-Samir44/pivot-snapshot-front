// Shared row/subtotal-grouping logic between PivotGrid.jsx (the live/on-screen table) and
// exportExcel.js (the .xlsx export) — extracted so both stay in sync automatically instead of
// maintaining two copies of the same subtotal/level-suppression rules.

// Mirrors the backend's HtmlSnapshotRenderer.ComputeVisibleLevels exactly — a level is only shown
// when it (or an earlier level) differs from the row directly above it, same as Excel's own
// "Tabular Form" pivot layout suppressing repeated parent labels.
export function computeVisibleLevels(labels, previousLabels) {
  const visible = new Array(labels.length).fill(false);
  let changed = previousLabels === null;
  for (let level = 0; level < labels.length; level++) {
    if (!changed && level < previousLabels.length && labels[level] === previousLabels[level]) {
      visible[level] = false;
    } else {
      changed = true;
      visible[level] = true;
    }
  }
  return visible;
}

// Real Excel's per-field "compact form" (PLAN 6 Phase 2, 2026-10-13) — groups consecutive row
// levels into one shared physical column wherever compactFlags chains them together ("compact
// form" pulls the NEXT field's labels into THIS field's own column). The old all-or-nothing
// Compact/Tabular split is just the degenerate all-true/all-false case of this same loop. Mirrors
// HtmlSnapshotRenderer.cs's BuildRowHeaderGroups exactly, so the live grid and a saved snapshot
// never disagree on which fields end up sharing a column. The last level's own flag is never
// consulted (nothing further to pull in).
export function buildRowHeaderGroups(rowDepth, compactFlags) {
  const groups = [];
  let i = 0;
  while (i < rowDepth) {
    const start = i;
    while (compactFlags[i] && i < rowDepth - 1) {
      i += 1;
    }
    groups.push({ start, end: i });
    i += 1;
  }
  return groups;
}

// A Total row's display caption and the row level it totals ("ER Total" at level 0, "Grand Total"
// at level -1). The backend sets caption/captionLevel (2026-10-14); a result saved BEFORE that (a
// snapshot's RawResultJson) has neither, so they're derived from the older label shape instead:
// SubtotalsProcessor's "Subtotal"/"Subtotal (Fn)" placeholders sit after the totaled level, and a
// Grand Total row has none.
export function totalCaptionInfo(row, measuresLevel = -1) {
  if (row.captionLevel !== undefined && row.captionLevel !== null) {
    return { level: row.captionLevel, caption: row.caption ?? "" };
  }
  const idx = row.labels.findIndex((l, i) => i !== measuresLevel && typeof l === "string" && l.startsWith("Subtotal"));
  if (idx > 0) {
    const fn = row.labels[idx].match(/^Subtotal \((\w+)\)$/)?.[1];
    return { level: idx - 1, caption: `${row.labels[idx - 1] || "(blank)"} ${fn ?? "Total"}` };
  }
  return { level: -1, caption: "Grand Total" };
}

// Turns the backend's flat row list into the rows real Excel actually DRAWS (2026-10-14, confirmed
// against the user-approved target shape). Shared by PivotGrid.jsx and exportExcel.js, mirrored by
// HtmlSnapshotRenderer.cs's BuildDisplayRows — keep all three in sync.
//
// - A "header level" is a row level merged into the same column as the next level (compact form).
//   Each of its items gets its OWN row (kind "header") above its children, like Excel's Compact
//   Form. A single Top subtotal for that item is absorbed into it (its values sit on the item row);
//   otherwise the header row has blank values (sourceIndex null).
// - A subtotal row (kind "subtotal") shows its caption ("ER Total"/"ER Sum") in the totaled level's
//   own column and leaves the other columns blank (only the Σ Values label is kept, so the row's
//   measure stays readable). A Top subtotal on a standalone (tabular) level, or several Top
//   functions on a header level, is moved to the BOTTOM of its item — Excel can't show those at top.
// - The Grand Total row (kind "grand") shows one caption in the first column.
// - Detail rows show each standalone level's label only when it changed from the row above (unless
//   Repeat Item Labels), and a merged column shows just its deepest level (its outer levels already
//   have their own header rows).
//
// drilldown: a PLAN 8 drilldown hierarchy keeps its own ragged, all-levels-in-one-column rendering —
// no header rows (its levels aren't separate fields).
//
// outline (2026-10-14): Excel's OUTLINE form (the report's Outline or Compact layout — Compact is
// outline form plus the merged column) gives EVERY outer level's items their own row, each in its
// own column, with the children starting on the next row. Only Tabular form puts an item on its
// first child's row. Not applied when Σ Values sits in Rows (that shape keeps the tabular rows).
//
// drilldownSubtotals (Drilldown v2, 2026-10-14): Excel's report Subtotals setting for an OLAP
// hierarchy. An expanded member's own row already carries its cube total — that's "Top". "Bottom"
// blanks that row's values and adds a "<member> Total" row (the SAME cube values, not a re-sum, so
// it's exact even for non-additive measures) after its descendants; "Off" just blanks it. A blanked
// parent keeps labelSourceIndex so its +/- toggle still knows which member it is.
//
// drillSlot (Drilldown v3, 2026-10-14): the Rows position of a drilldown hierarchy COMBINED with other
// fields (-1 = none; a drilldown ALONE still uses `drilldown`). Rows then have one label per field
// and row.drill = { level, depth, ancestorUniqueNames, memberUniqueName? } from the backend. The
// slot's labels are indented by that depth, and a member's BLOCK is its own rows plus every row with
// it among drill ancestors (same outer item) — so a total that belongs after a member (a Tabular
// Top moved to bottom, or an expanded parent's "Bottom" total when the slot is the last field)
// lands after the member's whole block, Hospital's rows included — real Excel's rule.
export function buildDisplayRows(rowHeaders, { rowHeaderGroups, rowFieldLayouts = [], repeatItemLabels = false, measuresLevel = -1, drilldown = false, outline = false, drilldownSubtotals = "Top", drillSlot = -1 }) {
  const groupOfLevel = [];
  rowHeaderGroups.forEach((g, gi) => {
    for (let l = g.start; l <= g.end; l++) groupOfLevel[l] = gi;
  });
  const depth = rowHeaders.reduce((max, r) => Math.max(max, r.labels.length), 0);
  const isHeaderLevel = (level) =>
    !drilldown &&
    groupOfLevel[level] !== undefined &&
    (level < rowHeaderGroups[groupOfLevel[level]].end || (outline && measuresLevel < 0 && level < depth - 1));
  const isMerged = (g) => g.start < g.end;
  const repeatFor = (level) => {
    const g = rowHeaderGroups[groupOfLevel[level]];
    return rowFieldLayouts[level]?.repeatItemLabels ?? (g && !isMerged(g) ? repeatItemLabels : false);
  };
  // Excel shows a member with an empty caption as "(blank)". The no-data placeholder (CubeQueryService
  // gives it an empty unique name too) stays truly blank. Display-only: the label itself stays "",
  // since filters turn captions back into MDX member references.
  const display = (text, uniqueName) => (text === "" && uniqueName !== "" ? "(blank)" : text);
  const pathOf = (labels, level) => labels.slice(0, level + 1);
  const belongs = (labels, path) => path.every((p, i) => labels[i] === p);
  const isGrand = (row) => row.isTotal && totalCaptionInfo(row, measuresLevel).level === -1;

  // Combined drilldown (drillSlot): the member a row stands for at the slot, its depth, the extra
  // indent the slot's label gets, and whether a row sits inside a member's block.
  const memberAt = (row) => row.drill?.memberUniqueName ?? row.uniqueNames?.[drillSlot];
  const indentFor = (g, l, row) => (isMerged(g) ? l - g.start : 0) + (drillSlot >= 0 && l === drillSlot ? (row?.drill?.depth ?? 0) : 0);
  const inDrillBlock = (row, outerLabels, member) =>
    drillSlot >= 0 && belongs(row.labels, outerLabels) && (memberAt(row) === member || (row.drill?.ancestorUniqueNames ?? []).includes(member));

  const out = [];
  const deferred = []; // stack of { path, level, indexes } — totals moved to the bottom of their item
  const emittedHeaders = new Set();
  let prevPath = null;

  const totalCells = (row, level, caption) =>
    rowHeaderGroups.map((g) => {
      const lines = [];
      const inGroup = level === -1 ? g.start === 0 : level >= g.start && level <= g.end;
      if (inGroup) {
        lines.push({ text: caption, indent: level === -1 ? 0 : indentFor(g, level, row), level: Math.max(level, 0) });
      } else if (measuresLevel > level && measuresLevel >= g.start && measuresLevel <= g.end) {
        lines.push({ text: row.labels[measuresLevel] ?? "", indent: isMerged(g) ? measuresLevel - g.start : 0, level: measuresLevel });
      }
      return { lines };
    });

  const pushSubtotal = (index, level) => {
    const row = rowHeaders[index];
    const { caption } = totalCaptionInfo(row, measuresLevel);
    out.push({ kind: "subtotal", sourceIndex: index, level, path: pathOf(row.labels, level), groupCells: totalCells(row, level, caption) });
  };

  // An outer label in its OWN column (outline form) still repeats on inner rows when its field has
  // Repeat Item Labels — like Excel; a label merged into a shared compact column never does.
  const showsOuterLabel = (g, l, visible) => (!isHeaderLevel(l) || (!isMerged(g) && repeatFor(l))) && (repeatFor(l) || visible[l]);

  // src = the row this header was derived from (for uniqueNames/drill depth). labelSourceIndex lets
  // a header at the drilldown slot keep its member's +/- toggle even with blank values.
  const pushHeader = (level, path, sourceIndex, src, srcIndex) => {
    const visible = computeVisibleLevels(path, prevPath);
    const groupCells = rowHeaderGroups.map((g) => {
      const lines = [];
      for (let l = g.start; l <= Math.min(g.end, level); l++) {
        if (l === level || showsOuterLabel(g, l, visible)) {
          lines.push({ text: display(path[l], src?.uniqueNames?.[l]), indent: indentFor(g, l, src), level: l });
        }
      }
      return { lines };
    });
    out.push({ kind: "header", sourceIndex, labelSourceIndex: srcIndex, level, path, groupCells });
    emittedHeaders.add(`${level}\u0001${path.join("\u0001")}`);
    prevPath = path;
  };

  // A deferred total at the drilldown slot stays pending through its member's WHOLE block (its
  // expanded descendants' rows included), not just rows sharing its own label.
  const stillInside = (row, entry) => belongs(row.labels, entry.path) || (entry.member !== undefined && inDrillBlock(row, entry.path.slice(0, drillSlot), entry.member));

  const flushDeferred = (row) => {
    while (deferred.length > 0 && (row === null || !stillInside(row, deferred[deferred.length - 1]))) {
      const d = deferred.pop();
      d.indexes.forEach((index) => pushSubtotal(index, d.level));
    }
  };

  // Drilldown "Bottom" subtotals: expanded parents whose "<member> Total" row is still pending,
  // innermost last — emitted once a row no longer belongs under them. A combined drilldown's parent
  // carries its member/outer labels, and its total goes only in the slot's own column.
  const openDrillParents = [];
  const flushDrillParents = (row) => {
    const keep = (p) => (p.member !== undefined ? inDrillBlock(row, p.outer, p.member) : belongs(row.labels, p.path));
    while (openDrillParents.length > 0 && (row === null || !keep(openDrillParents[openDrillParents.length - 1]))) {
      const p = openDrillParents.pop();
      const parent = rowHeaders[p.index];
      const caption = `${display(parent.labels[p.level], parent.uniqueNames?.[p.level])} Total`;
      out.push({
        kind: "subtotal",
        sourceIndex: p.index,
        level: p.level,
        path: parent.labels,
        groupCells: rowHeaderGroups.map((g) => ({
          lines: p.member === undefined || (p.level >= g.start && p.level <= g.end) ? [{ text: caption, indent: p.member === undefined ? p.level - g.start : indentFor(g, p.level, parent), level: p.level }] : [],
        })),
      });
    }
  };

  for (let i = 0; i < rowHeaders.length; i++) {
    const row = rowHeaders[i];

    if (isGrand(row)) {
      flushDrillParents(null);
      flushDeferred(null);
      const { caption } = totalCaptionInfo(row, measuresLevel);
      out.push({ kind: "grand", sourceIndex: i, level: -1, path: [], groupCells: totalCells(row, -1, caption) });
      prevPath = null;
      continue;
    }

    flushDrillParents(row);
    flushDeferred(row);

    if (row.isTotal) {
      const { level } = totalCaptionInfo(row, measuresLevel);
      const path = pathOf(row.labels, level);
      // A batch = every consecutive Total row for this same item and level (one per Custom
      // function). It's a TOP batch when the row right after it still belongs to the same item.
      let j = i;
      while (j < rowHeaders.length && rowHeaders[j].isTotal && !isGrand(rowHeaders[j]) && totalCaptionInfo(rowHeaders[j], measuresLevel).level === level && belongs(rowHeaders[j].labels, path)) {
        j += 1;
      }
      const indexes = Array.from({ length: j - i }, (_, k) => i + k);
      const after = rowHeaders[j];
      const isTop = after !== undefined && !isGrand(after) && belongs(after.labels, path);

      // At the drilldown slot the member's identity rides on the total itself (drill.memberUniqueName).
      const member = drillSlot >= 0 && level === drillSlot ? memberAt(row) : undefined;
      if (!isTop) {
        indexes.forEach((index) => pushSubtotal(index, level));
      } else if (isHeaderLevel(level) && indexes.length === 1) {
        pushHeader(level, path, i, row, member !== undefined ? i : undefined);
      } else {
        if (isHeaderLevel(level)) pushHeader(level, path, null, row, member !== undefined ? i : undefined);
        deferred.push({ path, level, indexes, member });
      }
      i = j - 1;
      continue;
    }

    // Detail row: first the item rows for any header level not shown yet.
    for (let level = 0; level < row.labels.length; level++) {
      if (isHeaderLevel(level) && !emittedHeaders.has(`${level}\u0001${pathOf(row.labels, level).join("\u0001")}`)) {
        pushHeader(level, pathOf(row.labels, level), null, row, drillSlot >= 0 && level === drillSlot ? i : undefined);
      }
    }

    // A no-data item (Show Items With No Data) whose only "child" is the empty placeholder: Excel
    // shows just the item row, so the item row takes this row's place instead of drawing a blank
    // child row under it.
    const last = out[out.length - 1];
    const innerIsPlaceholder = row.uniqueNames?.[row.labels.length - 1] === "";
    if (innerIsPlaceholder && last?.kind === "header" && last.sourceIndex === null && belongs(row.labels, last.path)) {
      last.sourceIndex = i;
      continue;
    }

    const visible = computeVisibleLevels(row.labels, prevPath);
    const groupCells = rowHeaderGroups.map((g) => {
      const lines = [];
      for (let l = g.start; l <= g.end && l < row.labels.length; l++) {
        if (showsOuterLabel(g, l, visible)) {
          lines.push({ text: display(row.labels[l], row.uniqueNames?.[l]), indent: indentFor(g, l, row), level: l });
        }
      }
      return { lines };
    });
    // An expanded drilldown member: its child comes right after it. Combined (drillSlot) only when
    // the slot is the LAST Rows field — with fields after it, the backend's own subtotal rows handle
    // the member's total (see SubtotalsProcessor.InsertDrillSlotSubtotals).
    const next = rowHeaders[i + 1];
    const isExpandedDrillParent = drilldown
      ? next !== undefined && !next.isTotal && next.labels.length > row.labels.length && belongs(next.labels, row.labels)
      : drillSlot >= 0 &&
        drillSlot === row.labels.length - 1 &&
        next !== undefined &&
        !next.isTotal &&
        belongs(next.labels, row.labels.slice(0, drillSlot)) &&
        (next.drill?.ancestorUniqueNames ?? []).includes(memberAt(row));
    if (isExpandedDrillParent && drilldownSubtotals !== "Top") {
      out.push({ kind: "detail", sourceIndex: null, labelSourceIndex: i, level: row.labels.length - 1, path: row.labels, groupCells });
      if (drilldownSubtotals === "Bottom") {
        openDrillParents.push(
          drilldown
            ? { index: i, level: row.labels.length - 1, path: row.labels }
            : { index: i, level: drillSlot, path: row.labels, outer: row.labels.slice(0, drillSlot), member: memberAt(row) },
        );
      }
    } else {
      out.push({ kind: "detail", sourceIndex: i, level: row.labels.length - 1, path: row.labels, groupCells });
    }
    prevPath = row.labels;
  }
  flushDrillParents(null);
  flushDeferred(null);

  // Insert Blank Line after each item (per field): after the LAST display row belonging to an item
  // of a flagged level — i.e. right after its own subtotal, when it has one.
  const blankLevels = rowFieldLayouts.map((fl, level) => (fl?.insertBlankLine ? level : null)).filter((l) => l !== null);
  out.forEach((r, idx) => {
    const next = out[idx + 1];
    r.blankLineAfter =
      next !== undefined &&
      blankLevels.some((level) => r.path.length > level && !(next.path.length > level && belongs(next.path, r.path.slice(0, level + 1))));
  });

  return out;
}

// Excel's +/- on an OUTERMOST item: collapsing it replaces all of that item's rows with ONE row
// carrying the item's own level-0 subtotal values. Only items that have such a subtotal can
// collapse (there's no value to show otherwise) — same rule as before this rewrite.
export function collapsibleKeys(displayRows) {
  const keys = new Set();
  displayRows.forEach((r) => {
    if (r.level === 0 && r.sourceIndex !== null && (r.kind === "subtotal" || r.kind === "header")) keys.add(r.path[0]);
  });
  return keys;
}

export function applyCollapse(displayRows, collapsedKeys, rowHeaderGroups) {
  if (collapsedKeys.size === 0) return displayRows;
  const out = [];
  for (let i = 0; i < displayRows.length; i++) {
    const key = displayRows[i].path[0];
    if (displayRows[i].path.length === 0 || !collapsedKeys.has(key)) {
      out.push(displayRows[i]);
      continue;
    }
    let j = i;
    let totalRow = null;
    while (j < displayRows.length && displayRows[j].path.length > 0 && displayRows[j].path[0] === key) {
      const r = displayRows[j];
      if (!totalRow && r.level === 0 && r.sourceIndex !== null && (r.kind === "subtotal" || r.kind === "header")) totalRow = r;
      j += 1;
    }
    out.push({
      kind: "collapsed",
      sourceIndex: totalRow?.sourceIndex ?? null,
      level: 0,
      path: [key],
      groupCells: rowHeaderGroups.map((g) => ({ lines: g.start === 0 ? [{ text: key === "" ? "(blank)" : key, indent: 0, level: 0 }] : [] })),
      blankLineAfter: displayRows[j - 1].blankLineAfter,
    });
    i = j - 1;
  }
  return out;
}

// Columns drilldown, Phase C2 (2026-10-14): the header rows real Excel draws above the value cells —
// mirrors the backend's ColumnHeaderLayout.Build exactly (keep both in sync). Returns one array per
// header row of { text, span, isTotal, sourceIndex, memberUniqueName }.
//
// Without a Columns drilldown (drillSlot -1): one row per Columns field, a Grand Total caption once in
// the top row. Since 2026-10-15 adjacent columns with the same path merge here too ("2024" drawn once,
// centered, over its months and its "2024 Total" column, like Excel), and a subtotal column
// (ColumnSubtotalsProcessor, captionLevel = its field's level) shows "2024 Total" on that field's row,
// blank below except the Σ Values row (measuresIndex, -1 when Σ isn't in Columns). With one, the slot's field becomes one row
// PER MEMBER DEPTH: ancestors above, the member at its own depth, blank below (a collapsed member or
// a total column). Adjacent columns with the same path down to a row merge into one cell there — an
// expanded parent spans its children, and its "Egypt Total" column (moved right of them by the
// backend's ColumnDrilldownLayoutProcessor) starts a new path at its depth.
export function buildColumnHeaderRows(columnHeaders, drillSlot = -1, measuresIndex = -1) {
  if (columnHeaders.length === 0) return [];
  const labelCount = columnHeaders[0].labels.length;
  const maxDepth = drillSlot >= 0 ? Math.max(...columnHeaders.map((c) => c.drill?.depth ?? 0)) : 0;
  const rowCount = labelCount + maxDepth;
  const entries = columnHeaders.map((column, c) => columnHeaderEntries(column, c, drillSlot, maxDepth, rowCount, measuresIndex));

  const rows = [];
  for (let r = 0; r < rowCount; r++) {
    const row = [];
    let c = 0;
    while (c < columnHeaders.length) {
      let end = c + 1;
      while (end < columnHeaders.length && entries[end][r].key === entries[c][r].key) end += 1;
      const e = entries[c][r];
      row.push({ text: e.text, span: end - c, isTotal: e.isTotal, sourceIndex: c, memberUniqueName: e.member });
      c = end;
    }
    rows.push(row);
  }
  return rows;
}

function columnHeaderEntries(column, c, drillSlot, maxDepth, rowCount, measuresIndex) {
  const entries = [];
  if (column.captionLevel === -1 && column.caption) {
    for (let r = 0; r < rowCount; r++) entries.push({ text: r === 0 ? column.caption : "", key: `grand:${c}:${r}`, isTotal: true, member: null });
    return entries;
  }
  if (drillSlot < 0) {
    // A subtotal column's level (ColumnSubtotalsProcessor); -1 for a detail column.
    const totalLevel = column.isTotal && typeof column.captionLevel === "number" && column.captionLevel >= 0 ? column.captionLevel : -1;
    let columnPath = "";
    for (let r = 0; r < rowCount; r++) {
      if (totalLevel < 0 || r < totalLevel) {
        columnPath += "\u0001" + column.labels[r];
        entries.push({ text: column.labels[r], key: columnPath, isTotal: Boolean(column.isTotal) && totalLevel < 0, member: null });
      } else if (r === totalLevel) {
        columnPath += "\u0001total:" + column.labels[r];
        entries.push({ text: column.caption ?? column.labels[r], key: columnPath, isTotal: true, member: null });
      } else {
        const label = r === measuresIndex ? column.labels[r] : "";
        columnPath += "\u0001" + label;
        entries.push({ text: label, key: columnPath, isTotal: true, member: null });
      }
    }
    return entries;
  }

  let path = "";
  const depth = column.drill?.depth ?? 0;
  const isTotal = Boolean(column.isTotal);
  const isSlotTotal = isTotal && column.captionLevel === drillSlot;
  let inTotal = false;
  column.labels.forEach((label, f) => {
    if (f !== drillSlot) {
      // By label, not unique name (see ColumnHeaderLayout.cs): a total column has no uniqueNames,
      // and its outer items must still merge with its siblings'.
      path += "\u0001" + label;
      entries.push({ text: label, key: path, isTotal: isTotal && (inTotal || f > drillSlot), member: null });
      return;
    }
    const ancestors = column.drill?.ancestorUniqueNames ?? [];
    const captions = column.drill?.ancestorCaptions ?? [];
    const member = isSlotTotal ? (column.drill?.memberUniqueName ?? null) : (column.uniqueNames?.[f] ?? null);
    for (let k = 0; k <= maxDepth; k++) {
      if (k < depth) {
        const ancestor = ancestors[k] ?? "";
        path += "\u0001" + ancestor;
        entries.push({ text: captions[k] ?? "", key: path, isTotal: false, member: ancestor });
      } else if (k === depth) {
        inTotal = isTotal;
        path += "\u0001" + (isSlotTotal ? "total:" : "") + (member ?? label);
        entries.push({ text: isSlotTotal ? (column.caption ?? label) : label, key: path, isTotal, member: isSlotTotal ? null : member });
      } else {
        path += "\u0001";
        entries.push({ text: "", key: path, isTotal, member: null });
      }
    }
  });
  return entries;
}
