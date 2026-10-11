// Mirrors the backend's PivotModels.MeasuresPlaceholder/MeasureIndexOf/BuildCaptionToIndex exactly
// — see that file's own WHY comment for the full story. "Σ Values" can now sit at ANY position
// within a request's rows/columns field list (not just "all the way in" on one or the other), so
// every place that used to assume measures were the fastest-varying factor on a fixed axis reads a
// header's own measure identity directly off its labels at the sentinel's position instead.
export const MEASURES_PLACEHOLDER = "[Measures]";

// Where Σ Values sits within the given axis field list, or -1 if it's on the OTHER axis entirely.
export function measuresIndexOf(axisFields) {
  return axisFields.indexOf(MEASURES_PLACEHOLDER);
}

// Mirrors the backend's PivotModels.NormalizeMeasuresPlacement exactly — a request with at least
// one Value but the sentinel in neither Rows nor Columns gets it appended to the end of Columns,
// same default the backend silently applies before ever building MDX.
//
// The backend being self-sufficient here was never the gap — MdxPivotQueryBuilder.Build calls its
// own NormalizeMeasuresPlacement regardless of caller discipline, so every query has always run
// correctly. What was missing is this exact mirror on the FRONTEND's own copy of `request` — Σ
// Values was built (PLAN 1) as a fully manual drag-and-drop pseudo-field with no default placement
// of its own, unlike the backend's safety net, so `request.rows`/`request.columns` stayed sentinel-
// free for the (overwhelmingly common) case of "just pick a measure, never bother dragging Σ
// Values anywhere yourself" — every frontend feature built on measuresIndexOf/measureIndexOf
// (the value-cell context menu, Number Format, Conditional Formatting, axis Sort-by-measure, the
// Columns/Rows drop-zone display itself) silently saw `measuresIndexOf(columns) === -1` in that
// case and degraded without any visible error — caught live, 2026-10-06, only because the new
// context menu's "no valueField, bail out" branch happens to fall back to the BROWSER's own context
// menu instead of ours, which is loud and obvious, unlike a silently-wrong number format would be.
export function normalizeMeasuresPlacement(request) {
  if (request.values.length === 0) return request;
  if (measuresIndexOf(request.rows) >= 0 || measuresIndexOf(request.columns) >= 0) return request;
  return { ...request, columns: [...request.columns, MEASURES_PLACEHOLDER] };
}

// Builds the caption->index lookup measureIndexOf needs, by reading it off the ACTUAL headers the
// backend already returned (not reconstructed from a value field's own name — the backend doesn't
// necessarily format captions identically to the raw field name). The first `valueCount` distinct
// captions encountered, in header order, are exactly the `valueCount` measures in the request's
// Values order — guaranteed by the backend always crossjoining measures in that same order at the
// sentinel's position.
export function buildCaptionToIndex(headersOnMeasuresAxis, measuresLevelIndex, valueCount) {
  const map = new Map();
  for (let i = 0; i < headersOnMeasuresAxis.length && map.size < valueCount; i++) {
    const caption = headersOnMeasuresAxis[i].labels[measuresLevelIndex];
    if (!map.has(caption)) {
      map.set(caption, map.size);
    }
  }
  return map;
}

// Maps a header back to which measure (by its 0-based index into the request's Values) it
// represents. Returns null when this axis doesn't carry Measures at all (measuresLevelIndex < 0).
export function measureIndexOf(header, measuresLevelIndex, captionToIndex) {
  if (measuresLevelIndex < 0) {
    return null;
  }
  const caption = header.labels[measuresLevelIndex];
  return captionToIndex.has(caption) ? captionToIndex.get(caption) : null;
}

// Real Excel's own behavior: the "Σ Values" breadcrumb only appears above the column headers when
// there are 2+ measures to distinguish between — with exactly one, whichever axis (if any) is on
// shows no separate "Σ Values" row/breadcrumb at all (caught live, 2026-10-06, "هي دي انت عاملها
// ليه انا مش لاقيها في الـ excel هل هي موجودة؟" against a one-measure table showing a "Σ Values"
// header above "Total Revenue"). Strips the sentinel out of the field-id list BEFORE it gets mapped
// to "Σ Values" text, rather than filtering the rendered string afterward, so this stays correct
// regardless of what that display label happens to be. Shared by the live builder (PivotBuilder.jsx)
// and the saved-snapshot interactive preview (SnapshotHistory.jsx) — previously only the live
// builder applied this, so a saved snapshot with one measure still showed the stray "Σ Values"
// header in its preview (caught live, 2026-10-07, "ليه الـ column دا؟؟؟؟").
//
// COLUMNS-ONLY for now: PivotGrid.jsx's Columns header renders columnFieldLabels as ONE joined
// string (`.join(" › ")`), with nothing else indexing into that array — safe to drop an entry from
// it. Rows renders rowFieldLabels[rowLevel] PER LEVEL, index-aligned against the real `rows` array
// (still containing the sentinel) for rowDepth/FieldHeaderMenu's own `field` prop — applying this
// same filter there would shift every label after the sentinel's position out of alignment with its
// actual field, corrupting every row header below it. Fixing that properly needs the rows side to
// look labels up BY FIELD ID instead of by raw array index; not done here since Σ Values defaults to
// the end of Columns (see this file's own normalizeMeasuresPlacement) and reaching Rows requires
// deliberately dragging it there — this fix covers what was actually reported, not a silently-
// incomplete claim of full symmetry.
export function hideMeasuresLabelIfSingleValue(fieldIds, valuesCount) {
  if (!fieldIds) return fieldIds;
  return valuesCount > 1 ? fieldIds : fieldIds.filter((id) => id !== MEASURES_PLACEHOLDER);
}
