// Shared row/subtotal-grouping logic between PivotGrid.jsx (the live/on-screen table) and
// exportExcel.js (the .xlsx export) — extracted so both stay in sync automatically instead of
// maintaining two copies of the same subtotal/level-suppression rules.

// The backend's SubtotalsProcessor inserts subtotals at EVERY row level except the innermost
// (e.g. Rows = [Specialty, Doctor, Month] gets both a per-Doctor and a nested per-Specialty
// subtotal), not just the outermost. An OUTER-level subtotal row is identifiable by every label
// after index 0 being the literal "Subtotal" placeholder (e.g. ["Cardiology", "Subtotal",
// "Subtotal"]); an INNER-level one (e.g. ["Cardiology", "Amin", "Subtotal"]) still starts with
// the same key but is NOT the terminator for the outer group — it's ordinary content inside it.
export function isOuterSubtotalFor(row, key) {
  return row.isTotal && row.labels[0] === key && row.labels.slice(1).every((l) => l === "Subtotal");
}

// Groups consecutive rows sharing the same first-level label into collapsible sections. A group
// is only collapsible when it ends in an OUTER-level Subtotal row — matches how Excel only offers
// +/- once subtotals exist.
export function computeGroups(rowHeaders) {
  const groups = [];
  let i = 0;
  while (i < rowHeaders.length) {
    if (rowHeaders[i].isTotal) {
      groups.push({ start: i, end: i + 1, key: null, hasSubtotal: false });
      i += 1;
      continue;
    }
    const start = i;
    const key = rowHeaders[i].labels[0];
    while (i < rowHeaders.length && rowHeaders[i].labels[0] === key && !isOuterSubtotalFor(rowHeaders[i], key)) {
      i += 1;
    }
    const hasSubtotal = i < rowHeaders.length && isOuterSubtotalFor(rowHeaders[i], key);
    const end = hasSubtotal ? i + 1 : i;
    if (hasSubtotal) i += 1;
    groups.push({ start, end, key, hasSubtotal });
  }
  return groups;
}

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

export function computeCompactVisibility(rowHeaders) {
  const result = [];
  let previousLabels = null;
  rowHeaders.forEach((row) => {
    result.push(computeVisibleLevels(row.labels, row.isTotal ? null : previousLabels));
    previousLabels = row.isTotal ? null : row.labels;
  });
  return result;
}
